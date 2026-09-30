"""
Tracker — private Notion-style personal workspace for Streamlit.

Storage:
- Preferred for Streamlit Community Cloud: Google Sheets via gspread.
- Local development fallback: JSON files in ./data.
- Excel remains available for backup/import/export.

Important:
- No database is required.
- No authentication is required.
- All records have stable UUIDs.
"""

from __future__ import annotations

import calendar
import datetime as dt
import hashlib
import html
import json
import os
import re
import tempfile
import threading
import uuid
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# Optional Google Sheets backend.
try:
    import gspread
    from google.oauth2.service_account import Credentials
except ImportError:
    gspread = None
    Credentials = None


# =============================================================================
# APP CONFIG
# =============================================================================

st.set_page_config(
    page_title="Tracker",
    page_icon="✅",
    layout="wide",
    initial_sidebar_state="expanded",
)

APP_DIR = Path(__file__).resolve().parent
LOCAL_DATA_DIR = Path(os.environ.get("TRACKER_DATA_DIR", APP_DIR / "data"))
LOCAL_DATA_DIR.mkdir(parents=True, exist_ok=True)

TODAY = dt.date.today()
NOW = dt.datetime.now()
esc = html.escape

AREAS = [
    "Personal", "Habit", "Work", "DSA", "Python", "Analytics",
    "Project", "Workout", "Class / Busy", "Random"
]
COLORS = dict(zip(AREAS, [
    "#7c5cff", "#22c55e", "#3b82f6", "#a855f7", "#eab308",
    "#f97316", "#ec4899", "#ef4444", "#64748b", "#14b8a6"
]))
STATUSES = ["To do", "Doing", "Done"]

SHEETS = {
    "Tasks": {
        "Task": ("text", None), "Area": ("sel", AREAS), "Subject": ("text", None),
        "Topic": ("text", None), "Sub-topic": ("text", None),
        "Status": ("sel", STATUSES), "Date": ("date", None),
        "Start": ("time", None), "End": ("time", None), "Deadline": ("date", None),
        "Minutes": ("num", None), "Project": ("text", None),
    },
    "Workout": {
        "Target": ("text", None), "Exercise": ("text", None), "Sets": ("num", None),
        "Reps": ("num", None), "Kg": ("num", None), "Minutes": ("num", None),
        "Date": ("date", None), "Notes": ("text", None),
    },
    "Wishlist": {
        "Item": ("text", None),
        "Category": ("sel", ["Want to do", "Want to visit", "Want to read", "Want to try", "Want to taste"]),
        "Status": ("sel", ["Wishlist", "Planned", "Done"]), "Notes": ("text", None),
    },
    "Jobs": {
        "Company": ("text", None), "Role": ("text", None),
        "Status": ("sel", ["Wishlist", "Applied", "Screening", "Interview", "Offer", "Rejected", "Ghosted"]),
        "Applied": ("date", None), "Follow-up": ("date", None),
        "Link": ("text", None), "Notes": ("text", None),
    },
    "Ideas": {
        "Idea": ("text", None), "Stage": ("sel", ["Raw", "Exploring", "Validating", "Building", "Shelved"]),
        "Problem": ("text", None), "Customer": ("text", None), "Next step": ("text", None),
    },
    "Projects": {
        "Project": ("text", None), "Status": ("sel", ["Idea", "Planning", "Active", "On hold", "Done"]),
        "Goal": ("text", None), "Deadline": ("date", None),
    },
    "Random": {
        "Dump": ("text", None), "Type": ("sel", ["Note", "Link", "Quote", "Snippet", "Question"]),
        "Details": ("text", None),
    },
    "Roadmap": {
        "Roadmap": ("text", None), "Heading": ("text", None), "Sub": ("text", None),
        "Topic": ("text", None), "Done": ("text", None),
    },
    "Habits": {"Habit": ("text", None), "Min": ("num", None)},
    "HabitLog": {"Date": ("text", None), "Habit": ("text", None), "Minutes": ("num", None)},
    "Targets": {"Target": ("text", None)},
    "Reminders": {"Text": ("text", None), "Done": ("text", None)},
    "WishNotes": {"Text": ("text", None)},
}

DEFAULTS = {
    "Habits": [
        {"Habit": "Wake up early", "Min": None},
        {"Habit": "Workout", "Min": 60},
        {"Habit": "DSA practice", "Min": 90},
        {"Habit": "Python backend", "Min": 60},
        {"Habit": "Data analytics", "Min": 60},
        {"Habit": "Read", "Min": 30},
        {"Habit": "Journal", "Min": None},
    ],
    "Targets": [
        {"Target": x} for x in
        ["Chest", "Back", "Biceps", "Triceps", "Shoulders", "Legs", "Core", "Swim", "Cardio"]
    ],
}

ICONS = {
    "Chest": "🏋️", "Back": "🧗", "Biceps": "💪", "Triceps": "🦾",
    "Shoulders": "🤸", "Legs": "🦵", "Core": "🎯", "Swim": "🏊", "Cardio": "🏃"
}


# =============================================================================
# STYLE
# =============================================================================

CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

html, body, [class*="css"], .stApp { font-family: Inter, system-ui, sans-serif; }
.stApp {
    background: radial-gradient(1100px 500px at 85% -10%, #2a1f55 0%, transparent 60%), #0e1015;
}
.block-container { padding-top: 1.6rem; max-width: 1500px; }
section[data-testid="stSidebar"] { background:#12141b; border-right:1px solid #232635; }
.brand {
    font-size:1.45rem; font-weight:700;
    background:linear-gradient(90deg,#9b87ff,#5eead4);
    -webkit-background-clip:text; color:transparent; margin-bottom:.6rem;
}
.hero h1 { font-size:2rem; margin:0; font-weight:700; }
.hero p { color:#8b90a3; margin:.1rem 0 1.1rem; }
.stat {
    background:linear-gradient(160deg,#1a1d29,#151823);
    border:1px solid #262a3b; border-radius:16px; padding:14px 16px;
}
.stat b { display:block; font-size:1.7rem; line-height:1.1; }
.stat span { color:#8b90a3; font-size:.8rem; text-transform:uppercase; letter-spacing:.06em; }
.stat i { display:block; height:5px; border-radius:9px; background:#252a3c; margin-top:8px; overflow:hidden; }
.stat i u { display:block; height:100%; background:linear-gradient(90deg,#7c5cff,#5eead4); }
div[data-testid="stVerticalBlockBorderWrapper"] {
    border-radius:16px!important; border-color:#262a3b!important; background:#151823;
}
.stButton>button, .stDownloadButton>button { border-radius:10px; border:1px solid #2c3146; font-weight:500; }
.stButton>button[kind="primary"] { background:linear-gradient(135deg,#7c5cff,#5b8cff); border:0; }
.tc { display:flex; gap:10px; align-items:center; padding:6px 0; }
.tc .dot { width:9px; height:9px; border-radius:50%; flex:none; }
.tc small { color:#8b90a3; }
.day { background:#151823; border:1px solid #262a3b; border-radius:14px; padding:12px; min-height:120px; }
.day h4 { margin:0 0 8px; font-size:.9rem; }
.day p { margin:4px 0; font-size:.82rem; border-left:3px solid var(--c); padding-left:7px; }
.tcard b { font-size:1.05rem; }
.login { text-align:center; margin-top:12vh; }
.login h1 { font-size:2.4rem; }
div[data-testid="stDialog"]>div { justify-content:flex-end!important; align-items:stretch!important; }
div[data-testid="stDialog"] div[role="dialog"] {
    position:fixed!important; right:0; top:0; margin:0!important;
    height:100vh!important; max-height:100vh!important; width:50vw!important;
    max-width:50vw!important; border-radius:20px 0 0 20px!important;
    background:#12141c;
}
@media(max-width:800px) {
    div[data-testid="stDialog"] div[role="dialog"] { width:100vw!important; max-width:100vw!important; }
}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


# =============================================================================
# WORKSPACE
# =============================================================================

# Single shared workspace: no login, passwords, or user accounts.
USER = "workspace"
NAME = "Tracker"


# =============================================================================
# STORAGE
# =============================================================================

LOCK = threading.RLock()
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

def use_google_sheets() -> bool:
    return (
        gspread is not None
        and Credentials is not None
        and "google_service_account" in st.secrets
        and "spreadsheet_id" in st.secrets
    )


def blank_record(schema: dict[str, tuple]) -> dict[str, Any]:
    return {c: None for c in schema}


def add_id(record: dict[str, Any]) -> dict[str, Any]:
    if not record.get("_id"):
        record["_id"] = str(uuid.uuid4())
    return record


def normalize_record(sheet: str, record: dict[str, Any]) -> dict[str, Any]:
    schema = SHEETS[sheet]
    out = {}
    for col, (kind, _) in schema.items():
        value = record.get(col)
        if value is None:
            out[col] = ""
        elif kind == "num":
            try:
                n = float(value)
                out[col] = int(n) if n.is_integer() else n
            except Exception:
                out[col] = ""
        elif kind == "date":
            if isinstance(value, dt.datetime):
                out[col] = value.date().isoformat()
            elif isinstance(value, dt.date):
                out[col] = value.isoformat()
            else:
                out[col] = str(value)
        elif kind == "time":
            if isinstance(value, dt.time):
                out[col] = value.strftime("%H:%M")
            else:
                out[col] = str(value)
        else:
            out[col] = str(value)
    out["_id"] = str(record.get("_id") or uuid.uuid4())
    return out


def validate_record(sheet: str, record: dict[str, Any]) -> tuple[bool, str]:
    schema = SHEETS[sheet]

    for col, (kind, options) in schema.items():
        value = record.get(col, "")

        if kind == "sel" and value not in ("", None) and value not in options:
            return False, f"{sheet}.{col}: invalid option {value!r}"

        if kind == "num" and value not in ("", None):
            try:
                if float(value) < 0:
                    return False, f"{sheet}.{col}: negative values are not allowed"
            except (ValueError, TypeError):
                return False, f"{sheet}.{col}: invalid number"

    if sheet == "Tasks":
        start, end = record.get("Start", ""), record.get("End", "")
        if start and end:
            try:
                s = dt.datetime.strptime(start, "%H:%M").time()
                e = dt.datetime.strptime(end, "%H:%M").time()
                if e <= s:
                    return False, "Tasks: End time must be after Start time"
            except ValueError:
                return False, "Tasks: invalid time"

    if sheet == "Jobs":
        link = str(record.get("Link", "")).strip()
        if link and not re.match(r"^https?://", link, re.I):
            return False, "Jobs: Link must start with http:// or https://"

    return True, ""


def local_file(user: str = USER) -> Path:
    return LOCAL_DATA_DIR / "tracker.json"


def local_load(user: str = USER) -> dict[str, list[dict[str, Any]]]:
    path = local_file(user)
    if not path.exists():
        return {sheet: [] for sheet in SHEETS}

    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
        db = {sheet: [] for sheet in SHEETS}
        for sheet in SHEETS:
            rows = raw.get(sheet, [])
            db[sheet] = [normalize_record(sheet, add_id(dict(r))) for r in rows]
        return db
    except Exception as exc:
        st.error(f"Could not read local data: {exc}")
        return {sheet: [] for sheet in SHEETS}


def local_save(user: str = USER, db: dict[str, list[dict[str, Any]]] | None = None) -> None:
    if db is None:
        db = {sheet: [] for sheet in SHEETS}
    path = local_file(user)
    payload = {
        sheet: [normalize_record(sheet, dict(r)) for r in rows]
        for sheet, rows in db.items()
    }
    # Atomic replacement: write a complete file, then swap it into place.
    fd, tmp_name = tempfile.mkstemp(prefix="tracker-", suffix=".json", dir=LOCAL_DATA_DIR)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp_name, path)
    finally:
        if os.path.exists(tmp_name):
            os.unlink(tmp_name)


def gs_client():
    info = dict(st.secrets["google_service_account"])
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)


def gs_book():
    return gs_client().open_by_key(str(st.secrets["spreadsheet_id"]))


def gs_load(user: str = USER) -> dict[str, list[dict[str, Any]]]:
    book = gs_book()
    db = {sheet: [] for sheet in SHEETS}

    for sheet, schema in SHEETS.items():
        try:
            ws = book.worksheet(sheet)
        except Exception:
            ws = book.add_worksheet(title=sheet, rows=100, cols=max(20, len(schema) + 1))
            ws.append_row(["_id"] + list(schema))
            continue

        values = ws.get_all_records()
        for row in values:
            db[sheet].append(normalize_record(sheet, row))

    return db


def gs_save(user: str = USER, db: dict[str, list[dict[str, Any]]] | None = None) -> None:
    if db is None:
        db = {sheet: [] for sheet in SHEETS}
    book = gs_book()

    for sheet, schema in SHEETS.items():
        title = sheet
        try:
            ws = book.worksheet(title)
        except Exception:
            ws = book.add_worksheet(title=title, rows=100, cols=max(20, len(schema) + 1))

        headers = ["_id"] + list(schema)
        rows = [headers]

        for record in db[sheet]:
            r = normalize_record(sheet, record)
            rows.append([r.get(h, "") for h in headers])

        ws.clear()
        # A single batch update prevents dozens of API calls per edit.
        ws.update("A1", rows, value_input_option="USER_ENTERED")


def seed_if_empty(db: dict[str, list[dict[str, Any]]]) -> bool:
    changed = False
    for sheet, defaults in DEFAULTS.items():
        if not db[sheet]:
            db[sheet] = [normalize_record(sheet, dict(r)) for r in defaults]
            changed = True
    return changed


def load_db() -> dict[str, list[dict[str, Any]]]:
    key = f"db"
    if key in st.session_state:
        return st.session_state[key]

    with LOCK:
        db = gs_load(USER) if use_google_sheets() else local_load(USER)

    if seed_if_empty(db):
        save_db(db)

    st.session_state[key] = db
    return db


def save_db(db: dict[str, list[dict[str, Any]]]) -> None:
    with LOCK:
        if use_google_sheets():
            gs_save(USER, db)
        else:
            local_save(USER, db)


DB = load_db()


# =============================================================================
# SERIALIZATION / EDITOR
# =============================================================================

def to_editor_df(sheet: str, rows: list[dict[str, Any]]) -> pd.DataFrame:
    schema = SHEETS[sheet]
    data = []

    for row in rows:
        r = normalize_record(sheet, row)
        item = {c: r.get(c, "") for c in schema}

        for c, (kind, _) in schema.items():
            if kind == "date":
                try:
                    item[c] = pd.to_datetime(item[c]).date() if item[c] else None
                except Exception:
                    item[c] = None
            elif kind == "time":
                try:
                    item[c] = dt.datetime.strptime(str(item[c]), "%H:%M").time() if item[c] else None
                except Exception:
                    item[c] = None
            elif kind == "num":
                if item[c] in ("", None):
                    item[c] = None
                else:
                    try:
                        item[c] = float(item[c])
                    except Exception:
                        item[c] = None

        data.append(item)

    # CRITICAL: always reset the index before passing data to st.data_editor.
    # This avoids several historical data_editor index/update bugs.
    return pd.DataFrame(data, columns=list(schema)).reset_index(drop=True)


def from_editor_df(sheet: str, df: pd.DataFrame) -> list[dict[str, Any]]:
    schema = SHEETS[sheet]
    result = []

    for _, row in df.reset_index(drop=True).iterrows():
        record = {}
        for c, (kind, _) in schema.items():
            value = row.get(c, "")

            if pd.isna(value):
                value = ""

            if kind == "date":
                value = "" if value == "" else (
                    value.isoformat() if isinstance(value, (dt.date, dt.datetime))
                    else str(value)
                )
            elif kind == "time":
                value = "" if value == "" else (
                    value.strftime("%H:%M") if isinstance(value, dt.time)
                    else str(value)[:5]
                )
            elif kind == "num":
                if value == "":
                    value = ""
                else:
                    try:
                        n = float(value)
                        value = int(n) if n.is_integer() else n
                    except Exception:
                        value = ""
            else:
                value = str(value)

            record[c] = value

        result.append(record)

    return result


def column_config_for(sheet: str, columns: list[str]) -> dict[str, Any]:
    schema = SHEETS[sheet]
    result = {}

    for c in columns:
        kind, options = schema[c]

        if kind == "sel":
            result[c] = st.column_config.SelectboxColumn(
                c, options=options, width="medium"
            )
        elif kind == "date":
            result[c] = st.column_config.DateColumn(c, width="medium")
        elif kind == "time":
            result[c] = st.column_config.TimeColumn(
                c, format="HH:mm", step=900, width="small"
            )
        elif kind == "num":
            result[c] = st.column_config.NumberColumn(
                c, min_value=0, width="small"
            )
        else:
            # Explicitly DO NOT set max_chars to a small value.
            # Streamlit's documented default is unlimited.
            result[c] = st.column_config.TextColumn(
                c,
                width="large",
                max_chars=None,
            )

    return result


def filtered_rows(sheet: str, mask=None) -> list[dict[str, Any]]:
    rows = DB[sheet]
    if mask is None:
        return rows

    df = to_editor_df(sheet, rows)
    keep = mask(df)
    ids = set(df.loc[keep, :].index)
    return [r for i, r in enumerate(rows) if i in ids]


def editor(
    sheet: str,
    mask=None,
    defaults: dict[str, Any] | None = None,
    show: list[str] | None = None,
    key: str | None = None,
) -> None:
    schema = SHEETS[sheet]
    full_df = to_editor_df(sheet, DB[sheet])

    if mask is None:
        visible = full_df.copy()
        hidden_df = full_df.iloc[0:0].copy()
        visible_original_indices = list(range(len(full_df)))
    else:
        keep = mask(full_df)
        visible = full_df.loc[keep].copy().reset_index(drop=True)
        hidden_df = full_df.loc[~keep].copy()
        visible_original_indices = list(full_df.index[keep])

    cols = show or list(schema)

    # If the table is filtered, preserve a stable editor key.
    editor_key = key or f"editor_{sheet}_{'_'.join(cols)}"

    edited = st.data_editor(
        visible[cols],
        num_rows="dynamic",
        width="stretch",
        hide_index=True,
        key=editor_key,
        column_config=column_config_for(sheet, cols),
    )

    edited = edited.reset_index(drop=True)

    # Restore hidden columns from the original visible rows.
    full_visible = pd.DataFrame(
        [dict(r) for r in visible.to_dict("records")],
        columns=cols
    )

    if defaults:
        for c, default in defaults.items():
            if c not in full_visible.columns:
                full_visible[c] = default
            else:
                full_visible[c] = full_visible[c].apply(
                    lambda x: default if x in ("", None) or pd.isna(x) else x
                )

    # Build visible records while preserving original IDs where possible.
    visible_records = []
    original_visible = visible.reset_index(drop=True)

    for i, row in edited.iterrows():
        rec = row.to_dict()

        if i < len(original_visible):
            original_row = original_visible.iloc[i].to_dict()
            rec["_id"] = DB[sheet][visible_original_indices[i]].get("_id")
        else:
            rec["_id"] = str(uuid.uuid4())

        for c in schema:
            if c not in rec:
                rec[c] = original_row.get(c, "") if i < len(original_visible) else ""

        rec = normalize_record(sheet, rec)
        valid, message = validate_record(sheet, rec)
        if not valid:
            st.error(message)
            return
        visible_records.append(rec)

    # Reconstruct the entire table from hidden rows + edited visible rows.
    hidden_records = []
    hidden_indices = set(full_df.index) - set(visible_original_indices)

    for original_i in sorted(hidden_indices):
        hidden_records.append(DB[sheet][original_i])

    combined = hidden_records + visible_records

    # For unfiltered editors, this is simply the edited table.
    if mask is None:
        combined = visible_records

    if serialize_records(sheet, DB[sheet]) != serialize_records(sheet, combined):
        DB[sheet] = combined
        save_db(DB)


def serialize_records(sheet: str, rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [
        normalize_record(sheet, r)
        for r in rows
    ]


# =============================================================================
# TASKS / QUICK ADD
# =============================================================================

def tasks_df() -> pd.DataFrame:
    return to_editor_df("Tasks", DB["Tasks"])


def add_task(task: dict[str, Any]) -> None:
    record = normalize_record("Tasks", task)
    valid, message = validate_record("Tasks", record)
    if not valid:
        st.error(message)
        return

    DB["Tasks"].append(record)
    save_db(DB)


def quick_add(day: dt.date, key: str) -> None:
    with st.form(f"quick_add_{key}", clear_on_submit=True):
        c = st.columns([3, 1.5, 1.5, 1.2, 1.2, 0.9])
        task = c[0].text_input("Task", placeholder="Write the full task...")
        area = c[1].selectbox("Area", AREAS)
        date_value = c[2].date_input("Day", day)
        start = c[3].time_input("Start", value=None, step=900)
        end = c[4].time_input("End", value=None, step=900)

        if c[5].form_submit_button("Add", type="primary", width="stretch"):
            if not task.strip():
                st.warning("Task cannot be empty.")
                return

            add_task({
                "Task": task.strip(),
                "Area": area,
                "Status": "To do",
                "Date": date_value.isoformat(),
                "Start": start.strftime("%H:%M") if start else "",
                "End": end.strftime("%H:%M") if end else "",
            })
            st.rerun()


def task_rows(df: pd.DataFrame, key: str) -> None:
    if df.empty:
        st.caption("Nothing here.")
        return

    for position, row in df.reset_index(drop=True).iterrows():
        record_id = str(row.get("_id", position))
        current = row.Status if row.Status in STATUSES else "To do"

        if pd.notna(row.Start) and row.Start:
            when = str(row.Start)[:5]
            if pd.notna(row.End) and row.End:
                when += f"–{str(row.End)[:5]}"
        else:
            when = "any time"

        extra = "".join(
            f" · {x}" for x in [row.Subject, row.Topic]
            if pd.notna(x) and str(x).strip()
        )

        c = st.columns([5, 3, .8], vertical_alignment="center")
        c[0].markdown(
            f'<div class="tc"><span class="dot" style="background:{COLORS.get(str(row.Area), "#7c5cff")}"></span>'
            f'<div><b>{esc(str(row.Task))}</b><br>'
            f'<small>{esc(str(row.Area))} · {esc(when)}{esc(extra)}</small></div></div>',
            unsafe_allow_html=True,
        )

        new_status = c[1].segmented_control(
            f"status_{key}_{record_id}",
            STATUSES,
            default=current,
            key=f"status_{key}_{record_id}",
            label_visibility="collapsed",
        )

        if new_status and new_status != current:
            for r in DB["Tasks"]:
                if r.get("_id") == record_id:
                    r["Status"] = new_status
                    break
            save_db(DB)
            st.rerun()

        if c[2].button("✕", key=f"delete_{key}_{record_id}"):
            DB["Tasks"] = [r for r in DB["Tasks"] if r.get("_id") != record_id]
            save_db(DB)
            st.rerun()


# =============================================================================
# PAGES
# =============================================================================

def stat(label: str, value: Any, sub: str = "", pct: int | None = None) -> str:
    bar = f'<i><u style="width:{max(0, min(100, pct))}%"></u></i>' if pct is not None else ""
    return (
        f'<div class="stat"><span>{esc(str(label))}</span><b>{esc(str(value))}</b>'
        f'<small style="color:#8b90a3">{esc(str(sub))}</small>{bar}</div>'
    )


def home() -> None:
    df = tasks_df()
    if df.empty:
        today_df = df
        overdue = df
    else:
        today_df = df[df["Date"] == TODAY]
        overdue = df[
            df["Date"].notna()
            & (df["Date"] < TODAY)
            & (df["Status"] != "Done")
        ]

    work = df[df["Area"].isin(["DSA", "Python", "Analytics", "Project", "Work"])] if not df.empty else df
    done_today = int((today_df["Status"] == "Done").sum()) if not today_df.empty else 0
    done_work = int((work["Status"] == "Done").sum()) if not work.empty else 0

    greeting = (
        "Good morning" if NOW.hour < 12
        else "Good afternoon" if NOW.hour < 18
        else "Good evening"
    )

    st.markdown(
        f'<div class="hero"><h1>{greeting}</h1>'
        f'<p>{TODAY:%A, %d %B %Y}</p></div>',
        unsafe_allow_html=True,
    )

    c = st.columns(4)
    c[0].markdown(
        stat("Tasks today", f"{done_today}/{len(today_df)}",
             pct=round(100 * done_today / len(today_df)) if len(today_df) else 0),
        unsafe_allow_html=True
    )
    c[1].markdown(
        stat("Habits today", f"{habit_day_pct(TODAY)}%", pct=habit_day_pct(TODAY)),
        unsafe_allow_html=True
    )
    c[2].markdown(
        stat("Work tasks", f"{done_work}/{len(work)}",
             pct=round(100 * done_work / len(work)) if len(work) else 0),
        unsafe_allow_html=True
    )
    c[3].markdown(stat("Overdue", len(overdue), "carried over"), unsafe_allow_html=True)

    L, R = st.columns([2.3, 1.1], gap="large")
    with L:
        with st.expander("➕ Add a task"):
            quick_add(TODAY, "home")
        st.markdown("##### Tasks for today")
        task_rows(today_df, "today")
        if len(overdue):
            with st.expander(f"⏳ Overdue ({len(overdue)})"):
                task_rows(overdue, "overdue")
    with R:
        pomodoro()
        today_habits()
        reminders()


def habit_names() -> list[str]:
    return [str(h["Habit"]) for h in DB["Habits"] if str(h.get("Habit", "")).strip()]


def habit_day_pct(day: dt.date) -> int:
    names = habit_names()
    if not names:
        return 0
    done = {
        r["Habit"] for r in DB["HabitLog"]
        if r.get("Date") == day.isoformat() and r.get("Habit") in names
    }
    return round(100 * len(done) / len(names))


def set_habit(day: dt.date, habit: str, completed: bool) -> None:
    date_key = day.isoformat()
    DB["HabitLog"] = [
        r for r in DB["HabitLog"]
        if not (r.get("Date") == date_key and r.get("Habit") == habit)
    ]
    if completed:
        # Store the minutes at completion time so historical data doesn't
        # change when the habit's planned minutes are edited later.
        configured = next(
            (r.get("Min") for r in DB["Habits"] if r.get("Habit") == habit), ""
        )
        DB["HabitLog"].append({
            "_id": str(uuid.uuid4()),
            "Date": date_key,
            "Habit": habit,
            "Minutes": configured if configured not in (None, "") else "",
        })
    save_db(DB)


def today_habits() -> None:
    with st.container(border=True):
        st.markdown(f"**✅ Today's habits** · {habit_day_pct(TODAY)}%")
        done = {
            r["Habit"] for r in DB["HabitLog"]
            if r.get("Date") == TODAY.isoformat()
        }

        for i, habit in enumerate(habit_names()):
            key = f"today_habit_{i}_{hashlib.sha1(habit.encode()).hexdigest()[:8]}"
            st.checkbox(
                habit,
                value=habit in done,
                key=key,
                on_change=set_habit,
                args=(TODAY, habit, not (habit in done)),
            )


def reminders() -> None:
    with st.container(border=True):
        st.markdown("**📝 Reminders**")
        with st.form("reminder_form", clear_on_submit=True):
            value = st.text_input(
                "Reminder",
                label_visibility="collapsed",
                placeholder="Short reminder",
            )
            if st.form_submit_button("Add", width="stretch") and value.strip():
                DB["Reminders"].append({
                    "_id": str(uuid.uuid4()),
                    "Text": value.strip(),
                    "Done": "",
                })
                save_db(DB)
                st.rerun()

        for r in list(DB["Reminders"]):
            rid = r["_id"]
            c1, c2 = st.columns([8, 1])
            c1.write(r["Text"])
            if c2.button("✓", key=f"rem_done_{rid}"):
                DB["Reminders"] = [x for x in DB["Reminders"] if x["_id"] != rid]
                save_db(DB)
                st.rerun()


def pomodoro() -> None:
    with st.container(border=True):
        st.markdown(
            """
            <div style="font-weight:600">🍅 Pomodoro
            <span id="pom_mode" style="color:#8b90a3;font-weight:400">· Focus</span></div>
            """,
            unsafe_allow_html=True,
        )
        components_html = """
        <div style="text-align:center;color:#e8eaf0;font-family:Inter,system-ui">
          <svg width="150" height="150" viewBox="0 0 120 120">
            <circle cx="60" cy="60" r="52" stroke="#252a3c" stroke-width="8" fill="none"/>
            <circle id="pom_ring" cx="60" cy="60" r="52" stroke="#7c5cff"
                    stroke-width="8" fill="none" stroke-linecap="round"
                    stroke-dasharray="326.7" stroke-dashoffset="0"
                    transform="rotate(-90 60 60)"/>
            <text id="pom_text" x="60" y="68" text-anchor="middle"
                  fill="#e8eaf0" font-size="24" font-weight="700">25:00</text>
          </svg>
          <div>
            <button onclick="pomStartPause()">Start / Pause</button>
            <button onclick="pomSet(1500,'Focus')">25</button>
            <button onclick="pomSet(300,'Break')">5</button>
            <button onclick="pomSet(900,'Long break')">15</button>
          </div>
        </div>
        <script>
        let pomTotal=1500, pomSeconds=1500, pomRunning=false, pomTimer=null;
        const pomText=document.getElementById("pom_text");
        const pomRing=document.getElementById("pom_ring");
        const pomMode=document.getElementById("pom_mode");
        function pomFormat(x) {
          return String(Math.floor(x/60)).padStart(2,"0")+":"+String(x%60).padStart(2,"0");
        }
        function pomRender() {
          pomText.textContent=pomFormat(pomSeconds);
          pomRing.style.strokeDashoffset=326.7*(1-pomSeconds/pomTotal);
        }
        function pomStartPause() {
          pomRunning=!pomRunning;
          clearInterval(pomTimer);
          if(pomRunning) {
            pomTimer=setInterval(()=>{
              if(pomSeconds>0) {
                pomSeconds--;
                pomRender();
              } else {
                pomRunning=false;
                clearInterval(pomTimer);
              }
            },1000);
          }
        }
        function pomSet(seconds,mode) {
          pomTotal=seconds; pomSeconds=seconds; pomMode.textContent="· "+mode;
          pomRunning=false; clearInterval(pomTimer); pomRender();
        }
        </script>
        """
        components.html(components_html, height=230)


def habits() -> None:
    st.title("✅ Habits")
    names = habit_names()

    if not names:
        st.info("Add habits in Habit settings below.")

    days = [TODAY + dt.timedelta(days=i) for i in range(-1, 6)]
    labels = [
        "Yesterday" if d < TODAY else "Today" if d == TODAY else d.strftime("%a %d %b")
        for d in days
    ]

    log = {
        (r.get("Date"), r.get("Habit"))
        for r in DB["HabitLog"]
    }

    rows = []
    for label, day in zip(labels, days):
        rows.append({
            "Day": label,
            **{name: (day.isoformat(), name) in log for name in names},
            "Done %": habit_day_pct(day),
        })

    ed = st.data_editor(
        pd.DataFrame(rows),
        hide_index=True,
        width="stretch",
        disabled=["Day", "Done %"],
        key=f"habits_grid",
        column_config={
            "Done %": st.column_config.ProgressColumn(
                "Done %", min_value=0, max_value=100, format="%d%%"
            ),
            **{
                name: st.column_config.CheckboxColumn(name)
                for name in names
            },
        },
    )

    current_log = {(r.get("Date"), r.get("Habit")): r for r in DB["HabitLog"]}
    changed = False

    for day, row in zip(days, ed.to_dict("records")):
        for name in names:
            key = (day.isoformat(), name)
            should_be_done = bool(row.get(name, False))
            is_done = key in current_log

            if should_be_done and not is_done:
                configured = next(
                    (r.get("Min") for r in DB["Habits"] if r.get("Habit") == name), ""
                )
                DB["HabitLog"].append(normalize_record("HabitLog", {
                    "Date": day.isoformat(),
                    "Habit": name,
                    "Minutes": configured if configured not in ("", None) else "",
                }))
                changed = True
            elif not should_be_done and is_done:
                DB["HabitLog"] = [
                    r for r in DB["HabitLog"]
                    if not (r.get("Date") == key[0] and r.get("Habit") == key[1])
                ]
                changed = True

    if changed:
        save_db(DB)

    st.markdown("#### Monthly overview")
    month = f"{TODAY:%Y-%m}"
    elapsed = TODAY.day

    overview = []
    for habit in DB["Habits"]:
        name = habit.get("Habit", "")
        if not name:
            continue

        completed = [
            r for r in DB["HabitLog"]
            if r.get("Habit") == name
            and str(r.get("Date", "")).startswith(month)
            and int(str(r.get("Date", "")).split("-")[-1] or 0) <= elapsed
        ]

        overview.append({
            "Habit": name,
            "Days done": len(completed),
            "Month % so far": round(100 * len(completed) / elapsed),
            "Minutes logged": sum(
                float(r.get("Minutes") or 0) for r in completed
            ),
        })

    st.dataframe(
        pd.DataFrame(overview),
        hide_index=True,
        width="stretch",
        column_config={
            "Month % so far": st.column_config.ProgressColumn(
                min_value=0, max_value=100, format="%d%%"
            )
        },
    )

    with st.expander("⚙️ Habit settings"):
        editor("Habits", key=f"habits_settings")


def tasks_page() -> None:
    st.title("📋 Tasks")
    with st.expander("➕ Quick add"):
        quick_add(TODAY, "tasks")

    areas = st.multiselect("Filter by area", AREAS, key=f"task_filter")
    mask = (lambda d: d["Area"].isin(areas)) if areas else None

    editor(
        "Tasks",
        mask=mask,
        key=f"tasks_editor_{'_'.join(areas)}",
        show=["Task", "Area", "Status", "Date", "Start", "End", "Deadline", "Minutes"],
    )


def roadmap_dialog(rid: str, title: str) -> None:
    rows = [
        (i, r) for i, r in enumerate(DB["Roadmap"])
        if r.get("Roadmap") == rid
    ]
    items = [(i, r) for i, r in rows if r.get("Topic")]
    done = sum(r.get("Done") == "1" for _, r in items)

    st.markdown(f"## 🗺️ {esc(title)}")
    st.progress(
        done / len(items) if items else 0.0,
        text=f"{done}/{len(items)} topics done",
    )

    headings = list(dict.fromkeys(
        r.get("Heading") for _, r in rows if r.get("Heading")
    ))

    for heading in headings:
        heading_items = [
            (i, r) for i, r in items if r.get("Heading") == heading
        ]
        heading_done = sum(r.get("Done") == "1" for _, r in heading_items)
        st.markdown(
            f"### {esc(str(heading))} "
            f"<small style='color:#8b90a3'>{heading_done}/{len(heading_items)}</small>",
            unsafe_allow_html=True,
        )

        subs = list(dict.fromkeys(
            r.get("Sub") for _, r in rows
            if r.get("Heading") == heading and r.get("Sub")
        ))

        for sub in [""] + subs:
            if sub:
                st.markdown(f"**{esc(str(sub))}**")

            for i, r in [
                (i, r) for i, r in heading_items if r.get("Sub", "") == sub
            ]:
                rid_key = r["_id"]
                checked = r.get("Done") == "1"
                value = st.checkbox(
                    str(r.get("Topic", "")),
                    value=checked,
                    key=f"roadmap_topic_{rid_key}",
                )
                if value != checked:
                    r["Done"] = "1" if value else ""
                    save_db(DB)
                    st.rerun()

    with st.expander("✏️ Edit structure", expanded=not rows):
        headings = headings or ["Week 1"]

        heading = st.selectbox(
            "Heading", headings, key=f"roadmap_heading_{rid}"
        )
        new_heading = st.text_input(
            "New heading", key=f"roadmap_new_heading_{rid}"
        )
        if st.button("Add heading", key=f"roadmap_add_heading_{rid}") and new_heading.strip():
            DB["Roadmap"].append(normalize_record("Roadmap", {
                "Roadmap": rid, "Heading": new_heading.strip(), "Sub": "",
                "Topic": "", "Done": ""
            }))
            save_db(DB)
            st.rerun()

        subs = list(dict.fromkeys(
            r.get("Sub") for _, r in rows
            if r.get("Heading") == heading and r.get("Sub")
        ))
        sub = st.selectbox(
            "Sub-heading", ["(none)"] + subs,
            key=f"roadmap_sub_{rid}"
        )
        topics = st.text_area(
            "Topics, one per line",
            key=f"roadmap_topics_{rid}",
            height=120,
        )
        if st.button("Add topics", key=f"roadmap_add_topics_{rid}", type="primary"):
            new_rows = []
            for topic in topics.splitlines():
                topic = topic.strip()
                if topic:
                    new_rows.append(normalize_record("Roadmap", {
                        "Roadmap": rid,
                        "Heading": heading,
                        "Sub": "" if sub == "(none)" else sub,
                        "Topic": topic,
                        "Done": "",
                    }))
            if new_rows:
                DB["Roadmap"].extend(new_rows)
                save_db(DB)
                st.rerun()

        if st.button(
            f'🗑️ Delete heading "{heading}" and everything in it',
            key=f"roadmap_delete_{rid}",
        ):
            DB["Roadmap"] = [
                r for r in DB["Roadmap"]
                if not (r.get("Roadmap") == rid and r.get("Heading") == heading)
            ]
            save_db(DB)
            st.rerun()


def study(area: str, rid: str, title: str) -> None:
    items = [
        r for r in DB["Roadmap"]
        if r.get("Roadmap") == rid and r.get("Topic")
    ]
    done = sum(r.get("Done") == "1" for r in items)

    with st.container(border=True):
        c = st.columns([3, 1], vertical_alignment="center")
        c[0].markdown(
            f"**🗺️ {esc(title)} roadmap**  \n"
            f"<small style='color:#8b90a3'>{done}/{len(items)} topics done</small>",
            unsafe_allow_html=True,
        )
        if c[1].button(
            "Open roadmap",
            key=f"open_roadmap_{rid}",
            type="primary",
            width="stretch",
        ):
            roadmap_dialog(rid, title)

    st.caption(
        "Subject = section header. Each task has Topic and Sub-topic. "
        "Deadline is optional."
    )
    editor(
        "Tasks",
        mask=lambda d: d["Area"] == area,
        defaults={"Area": area, "Status": "To do"},
        key=f"study_{rid}",
        show=[
            "Task", "Subject", "Topic", "Sub-topic", "Status",
            "Deadline", "Minutes", "Date", "Start", "End"
        ],
    )


def work() -> None:
    st.title("💼 Work")
    tabs = st.tabs([
        "DSA", "Python Backend", "Data Analytics",
        "Ideas", "Projects", "Random"
    ])

    with tabs[0]:
        study("DSA", "dsa", "DSA")
    with tabs[1]:
        study("Python", "py", "Python Backend")
    with tabs[2]:
        study("Analytics", "da", "Data Analytics")
    with tabs[3]:
        editor("Ideas", key=f"ideas")
    with tabs[4]:
        editor("Projects", key=f"projects")
        st.markdown("##### Project tasks")
        editor(
            "Tasks",
            mask=lambda d: d["Area"] == "Project",
            defaults={"Area": "Project", "Status": "To do"},
            key=f"project_tasks",
            show=[
                "Task", "Project", "Status",
                "Deadline", "Minutes", "Date", "Start", "End"
            ],
        )
    with tabs[5]:
        editor("Random", key=f"random")


def workout() -> None:
    st.title("💪 Workout")
    targets = [
        r["Target"] for r in DB["Targets"]
        if str(r.get("Target", "")).strip()
    ]
    counts = {}
    for r in DB["Workout"]:
        counts[r.get("Target")] = counts.get(r.get("Target"), 0) + 1

    selected = st.session_state.get("workout_target")
    if selected not in targets:
        selected = None

    cols = st.columns(4)
    for i, target in enumerate(targets):
        with cols[i % 4].container(border=True):
            st.markdown(
                f'<div class="tcard"><span style="font-size:1.8rem">'
                f'{ICONS.get(target, "🔥")}</span><br><b>{esc(target)}</b><br>'
                f'<small style="color:#8b90a3">{counts.get(target, 0)} entries</small></div>',
                unsafe_allow_html=True,
            )
            st.button(
                "Open",
                key=f"workout_open_{hashlib.sha1(target.encode()).hexdigest()[:10]}",
                width="stretch",
                type="primary" if selected == target else "secondary",
                on_click=lambda target=target: st.session_state.update(
                    workout_target=target
                ),
            )

    with st.expander("➕ Add target area"):
        with st.form("target_form", clear_on_submit=True):
            value = st.text_input("Name", placeholder="e.g. Forearms")
            if st.form_submit_button("Add") and value.strip():
                if value.strip() not in targets:
                    DB["Targets"].append(normalize_record(
                        "Targets", {"Target": value.strip()}
                    ))
                    save_db(DB)
                    st.rerun()
                else:
                    st.warning("That target already exists.")

    if selected:
        st.markdown(f"### {ICONS.get(selected, '🔥')} {esc(selected)}")
        editor(
            "Workout",
            mask=lambda d: d["Target"] == selected,
            defaults={"Target": selected},
            key=f"workout_editor_{hashlib.sha1(selected.encode()).hexdigest()[:10]}",
            show=["Exercise", "Sets", "Reps", "Kg", "Minutes", "Date", "Notes"],
        )
    else:
        st.caption("Click a target card to open its table.")


def wishlist() -> None:
    st.title("⭐ Wishlist")
    current = (DB["WishNotes"][0].get("Text", "") if DB["WishNotes"] else "")
    value = st.text_area(
        "Notes",
        current,
        height=110,
        placeholder="Quick notes: things you want to add later…",
    )
    if value != current:
        DB["WishNotes"] = [{
            "_id": DB["WishNotes"][0]["_id"] if DB["WishNotes"] else str(uuid.uuid4()),
            "Text": value,
        }]
        save_db(DB)

    editor("Wishlist", key=f"wishlist")


def jobs() -> None:
    st.title("🎯 Job Tracker")
    df = to_editor_df("Jobs", DB["Jobs"])
    options = SHEETS["Jobs"]["Status"][1]

    for c, option in zip(st.columns(len(options)), options):
        count = int((df["Status"] == option).sum()) if not df.empty else 0
        c.markdown(stat(option, count), unsafe_allow_html=True)

    st.write("")
    editor("Jobs", key=f"jobs")


def week() -> None:
    st.title("🗓️ Plan: Week")
    with st.expander("➕ Quick add"):
        quick_add(TODAY, "week")

    df = tasks_df()
    cols = st.columns(4)

    for i in range(7):
        day = TODAY + dt.timedelta(days=i)
        if df.empty:
            x = df
        else:
            x = df[
                (df["Date"] == day)
                & (df["Status"] != "Done")
            ].sort_values("Start", na_position="last")

        body_parts = []
        for _, r in x.iterrows():
            time_text = str(r["Start"])[:5] if pd.notna(r["Start"]) and r["Start"] else "any time"
            body_parts.append(
                f'<p style="--c:{COLORS.get(str(r["Area"]), "#7c5cff")}">'
                f'<b>{esc(str(r["Task"]))}</b><br>'
                f'<small style="color:#8b90a3">{esc(time_text)}</small></p>'
            )

        body = "".join(body_parts) or '<small style="color:#6b7087">Free</small>'
        cols[i % 4].markdown(
            f'<div class="day"><h4>{day:%a %d %b}</h4>{body}</div>',
            unsafe_allow_html=True,
        )


def export_excel() -> bytes:
    from io import BytesIO

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet, schema in SHEETS.items():
            rows = [normalize_record(sheet, r) for r in DB[sheet]]
            df = pd.DataFrame(rows)
            columns = ["_id"] + list(schema)
            df.reindex(columns=columns).to_excel(
                writer, sheet_name=sheet[:31], index=False
            )
    return output.getvalue()


# =============================================================================
# NAVIGATION
# =============================================================================

PAGES = {
    "🏠 Home": home,
    "✅ Habits": habits,
    "📋 Tasks": tasks_page,
    "💪 Workout": workout,
    "💼 Work": work,
    "⭐ Wishlist": wishlist,
    "🎯 Job Tracker": jobs,
    "🗓️ Plan: Week": week,
}

with st.sidebar:
    st.markdown('<div class="brand">✅ Tracker</div>', unsafe_allow_html=True)
    page = st.radio(
        "Go to",
        list(PAGES),
        key=f"page",
        label_visibility="collapsed",
    )
    st.divider()

    st.download_button(
        "⬇️ Download my Excel backup",
        data=export_excel(),
        file_name=f"tracker.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
    )

    st.info(
        "Storage: Google Sheets"
        if use_google_sheets()
        else "Storage: local JSON (safe for local development; configure Google Sheets for persistent cloud hosting)."
    )

PAGES[page]()
