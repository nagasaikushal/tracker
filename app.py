"""
Tracker — private Notion-style personal workspace for Streamlit.

Storage:
- Supabase PostgreSQL via the Supabase Python client.
- Excel remains available for backup/import/export.

Important:
- Supabase is required for this version.
- Three-role authentication: admin, kushal, jahnavi.
- Admin records are shared with both users; user records are private to their owner.
- All records have stable UUIDs.
"""

from __future__ import annotations

import datetime as dt
import hashlib
import hmac
import html
import math
import os
import re
import threading
import time
import uuid
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st

# Supabase backend.
try:
    from supabase import create_client, Client
except ImportError:
    create_client = None
    Client = Any


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
        "CompletedAt": ("text", None),
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
        "Order": ("num", None),
    },
    "Habits": {"Habit": ("text", None), "Min": ("num", None)},
    "HabitLog": {"Date": ("text", None), "Habit": ("text", None), "Minutes": ("num", None)},
    "Targets": {"Target": ("text", None)},
    "Reminders": {"Text": ("text", None), "Done": ("text", None), "DoneAt": ("text", None)},
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
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@500;600;700;800&family=Inter:wght@400;500;600;700;800&display=swap');

:root{
  --bg:#08090d;
  --bg-2:#0c0e14;
  --surface:#10131b;
  --surface-2:#141822;
  --surface-3:#181c27;
  --surface-hover:#1c2130;
  --ink:#f4f6fb;
  --ink-2:#d9deea;
  --muted:#8992a7;
  --muted-2:#657087;
  --line:#252b39;
  --line-2:#303749;

  --purple:#8b7cff;
  --purple-2:#b39cff;
  --blue:#55a7ff;
  --cyan:#46d8e8;
  --mint:#42d6a4;
  --yellow:#ffc857;
  --coral:#ff7185;
  --pink:#ee75c9;

  --glow-purple:rgba(139,124,255,.18);
  --glow-blue:rgba(85,167,255,.13);
  --glow-mint:rgba(66,214,164,.12);

  --shadow:0 20px 60px rgba(0,0,0,.32);
  --shadow-sm:0 8px 26px rgba(0,0,0,.24);
  --radius:18px;
}

/* =========================
   GLOBAL / CANVAS
   ========================= */
html,body,[class*="css"],.stApp{
  font-family:Inter,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif!important;
  color:var(--ink)!important;
}
.stApp{
  min-height:100vh;
  background:
    radial-gradient(900px 500px at 82% -8%,rgba(139,124,255,.14),transparent 60%),
    radial-gradient(750px 420px at -8% 18%,rgba(70,216,232,.07),transparent 58%),
    radial-gradient(700px 450px at 55% 110%,rgba(66,214,164,.055),transparent 62%),
    var(--bg)!important;
}
.stApp:before{
  content:"";
  position:fixed;
  inset:0;
  pointer-events:none;
  opacity:.22;
  background-image:
    linear-gradient(rgba(255,255,255,.025) 1px,transparent 1px),
    linear-gradient(90deg,rgba(255,255,255,.025) 1px,transparent 1px);
  background-size:48px 48px;
  mask-image:linear-gradient(to bottom,black,transparent 88%);
  z-index:0;
}
.block-container{
  position:relative;
  z-index:1;
  max-width:1500px;
  padding:2.2rem 2.5rem 5rem!important;
}
#MainMenu{visibility:hidden}
.stDeployButton{display:none}
.stAppHeader{
  background:rgba(8,9,13,.72)!important;
  backdrop-filter:blur(20px)!important;
  border-bottom:1px solid rgba(255,255,255,.035)!important;
}
header[data-testid="stHeader"]{height:2.5rem}
footer{visibility:hidden}

/* Scrollbar */
::-webkit-scrollbar{width:9px;height:9px}
::-webkit-scrollbar-track{background:#090a0f}
::-webkit-scrollbar-thumb{background:#2b3140;border-radius:99px;border:2px solid #090a0f}
::-webkit-scrollbar-thumb:hover{background:#454d61}

/* =========================
   SIDEBAR / NAVIGATION
   ========================= */
section[data-testid="stSidebar"]{
  background:
    linear-gradient(180deg,rgba(16,19,27,.98),rgba(10,12,17,.98))!important;
  border-right:1px solid #202532!important;
  box-shadow:18px 0 60px rgba(0,0,0,.28)!important;
}
section[data-testid="stSidebar"]>div{
  padding:1rem .78rem 1.2rem!important;
}
.brand{
  display:flex;
  align-items:center;
  gap:11px;
  padding:.55rem .62rem 1.2rem;
  font-family:"DM Sans",Inter,sans-serif;
  font-size:1.3rem;
  font-weight:800;
  color:var(--ink)!important;
  letter-spacing:-.035em;
}
.brand-mark{
  display:inline-grid;
  place-items:center;
  width:39px;height:39px;
  border-radius:13px;
  background:linear-gradient(135deg,#7868ff,#b09aff);
  color:white;
  box-shadow:0 0 28px rgba(139,124,255,.3);
}
section[data-testid="stSidebar"] [data-testid="stRadio"]>label{display:none}
section[data-testid="stSidebar"] [role="radiogroup"]{gap:5px}
section[data-testid="stSidebar"] [role="radio"]{
  border:1px solid transparent;
  border-radius:13px;
  padding:10px 11px!important;
  color:#8992a7!important;
  font-weight:600;
  transition:all .18s ease;
}
section[data-testid="stSidebar"] [role="radio"]:hover{
  background:#171b25;
  border-color:#252b39;
  color:#e7eaf2!important;
  transform:translateX(2px);
}
section[data-testid="stSidebar"] [role="radio"][aria-checked="true"]{
  background:linear-gradient(100deg,rgba(139,124,255,.18),rgba(139,124,255,.06))!important;
  border-color:rgba(139,124,255,.28)!important;
  color:#c9c1ff!important;
  box-shadow:inset 3px 0 0 var(--purple),0 8px 24px rgba(0,0,0,.16);
}
.sidebar-user{
  padding:13px;
  margin:13px 2px;
  border:1px solid #262c3a;
  background:linear-gradient(145deg,#141822,#10131b);
  border-radius:16px;
  box-shadow:inset 0 1px rgba(255,255,255,.025);
}
.sidebar-avatar{
  display:inline-grid;place-items:center;
  width:36px;height:36px;border-radius:50%;
  background:linear-gradient(135deg,#302b65,#211d42);
  color:#bdb5ff;font-weight:800;
  margin-right:8px;vertical-align:middle;
  border:1px solid #4a427c;
}
.sidebar-role{font-size:.73rem;color:var(--muted);margin-top:5px;line-height:1.4}

/* =========================
   TYPOGRAPHY
   ========================= */
h1,h2,h3,h4,h5,h6{
  font-family:"DM Sans",Inter,sans-serif!important;
  color:var(--ink)!important;
  letter-spacing:-.035em!important;
}
.stMarkdown h1{font-size:2.25rem!important;margin:.15rem 0 .4rem!important;font-weight:800!important}
.stMarkdown h2{font-size:1.5rem!important;margin:1.5rem 0 .7rem!important;font-weight:750!important}
.stMarkdown h3{font-size:1.12rem!important;margin:1.1rem 0 .55rem!important;font-weight:700!important}
.stCaption,.stMarkdown small{color:var(--muted)!important}
p,li{color:var(--ink-2)}
a{color:#a99dff!important}

/* =========================
   BUTTONS
   ========================= */
.stButton>button,.stDownloadButton>button{
  min-height:40px!important;
  border-radius:12px!important;
  border:1px solid #2b3140!important;
  background:linear-gradient(180deg,#181c26,#13161e)!important;
  color:#e8ebf2!important;
  font-weight:650!important;
  box-shadow:0 7px 20px rgba(0,0,0,.18)!important;
  transition:all .18s ease!important;
}
.stButton>button:hover,.stDownloadButton>button:hover{
  border-color:#4a426f!important;
  color:#fff!important;
  background:linear-gradient(180deg,#202434,#171b26)!important;
  transform:translateY(-1px);
  box-shadow:0 10px 28px rgba(0,0,0,.28),0 0 20px rgba(139,124,255,.08)!important;
}
.stButton>button[kind="primary"]{
  background:linear-gradient(135deg,#7767ff,#9a87ff)!important;
  color:#fff!important;
  border:0!important;
  box-shadow:0 9px 28px rgba(123,105,255,.28),inset 0 1px rgba(255,255,255,.22)!important;
}
.stButton>button[kind="primary"]:hover{
  background:linear-gradient(135deg,#8879ff,#a392ff)!important;
  box-shadow:0 12px 34px rgba(123,105,255,.38)!important;
}

/* =========================
   INPUTS / SELECTS
   ========================= */
.stTextInput input,.stTextArea textarea,.stNumberInput input,
.stDateInput input,.stTimeInput input,
.stSelectbox div[data-baseweb="select"]>div,
.stMultiSelect div[data-baseweb="select"]>div{
  background:#11151d!important;
  color:#edf0f7!important;
  border:1px solid #2a3040!important;
  border-radius:12px!important;
  box-shadow:inset 0 1px rgba(255,255,255,.025)!important;
}
.stTextInput input:focus,.stTextArea textarea:focus,.stNumberInput input:focus,
.stDateInput input:focus,.stTimeInput input:focus{
  border-color:#7163d5!important;
  box-shadow:0 0 0 3px rgba(139,124,255,.12),0 0 22px rgba(139,124,255,.05)!important;
}
.stSelectbox label,.stMultiSelect label,.stTextInput label,.stTextArea label,
.stNumberInput label,.stDateInput label,.stTimeInput label{
  color:#aab2c2!important;
  font-weight:600!important;
}
div[data-baseweb="popover"],div[data-baseweb="menu"]{
  background:#141822!important;
  border:1px solid #303749!important;
  box-shadow:0 20px 50px rgba(0,0,0,.5)!important;
}
div[data-baseweb="menu"] *{color:#e9ecf3!important}
div[data-baseweb="menu"] [aria-selected="true"]{background:#25213f!important}
input::placeholder,textarea::placeholder{color:#5f687c!important}

/* =========================
   CONTAINERS / CARDS
   ========================= */
div[data-testid="stVerticalBlockBorderWrapper"]{
  border:1px solid #252b39!important;
  border-radius:18px!important;
  background:linear-gradient(145deg,rgba(19,23,32,.96),rgba(14,17,24,.96))!important;
  box-shadow:var(--shadow-sm)!important;
}
.hero{
  position:relative;
  padding:28px 30px;
  margin:0 0 22px;
  min-height:145px;
  border:1px solid #302b54;
  border-radius:24px;
  background:
    radial-gradient(420px 180px at 86% 0%,rgba(139,124,255,.22),transparent 65%),
    radial-gradient(300px 150px at 20% 100%,rgba(70,216,232,.08),transparent 70%),
    linear-gradient(135deg,#151827,#0e1118);
  box-shadow:0 22px 60px rgba(0,0,0,.3),inset 0 1px rgba(255,255,255,.045);
  overflow:hidden;
}
.hero:before{
  content:"";
  position:absolute;
  width:320px;height:320px;
  right:-160px;top:-190px;
  border-radius:50%;
  border:1px solid rgba(179,156,255,.14);
  box-shadow:0 0 80px rgba(139,124,255,.1);
}
.hero:after{
  content:"";
  position:absolute;
  width:7px;height:70px;
  left:0;top:35px;
  border-radius:0 99px 99px 0;
  background:linear-gradient(#9b89ff,#54d7e7);
  box-shadow:0 0 24px rgba(139,124,255,.45);
}
.hero h1{font-size:2.35rem!important;margin:0!important;position:relative;z-index:1}
.hero p{color:#8f98ad;margin:.45rem 0 0;position:relative;z-index:1}

/* Stats */
.stat{
  position:relative;
  background:linear-gradient(145deg,#151923,#10131a);
  border:1px solid #252b39;
  border-radius:18px;
  padding:18px 18px 16px;
  box-shadow:var(--shadow-sm);
  overflow:hidden;
  min-height:112px;
  transition:.2s ease;
}
.stat:hover{transform:translateY(-2px);border-color:#383f51;box-shadow:0 16px 36px rgba(0,0,0,.3)}
.stat:before{
  content:"";position:absolute;left:0;top:0;width:4px;height:100%;
  background:linear-gradient(#8b7cff,#55d8e8);
}
.stat b{display:block;font-size:1.72rem;line-height:1.1;margin:.25rem 0 .2rem;color:#f5f6fa;font-family:"DM Sans",Inter,sans-serif}
.stat span{color:#7f899d;font-size:.7rem;text-transform:uppercase;letter-spacing:.09em;font-weight:800}
.stat i{display:block;height:6px;border-radius:99px;background:#252a36;margin-top:11px;overflow:hidden}
.stat i u{display:block;height:100%;background:linear-gradient(90deg,#8b7cff,#48d7d1);border-radius:99px}
.tc{display:flex;gap:10px;align-items:center;padding:8px 0}
.tc .dot{width:9px;height:9px;border-radius:50%;flex:none;box-shadow:0 0 10px currentColor}
.tc small{color:#8992a7}
.day{background:linear-gradient(145deg,#151923,#10131a);border:1px solid #252b39;border-radius:16px;padding:14px;min-height:125px;box-shadow:var(--shadow-sm)}
.day h4{margin:0 0 9px;font-size:.9rem!important}
.day p{margin:5px 0;font-size:.82rem;border-left:3px solid var(--c);padding-left:8px;color:#c8ceda}
.tcard b{font-size:1.05rem}

/* =========================
   EXPANDERS
   ========================= */
details[data-testid="stExpander"]{
  border:1px solid #282e3c!important;
  border-radius:15px!important;
  background:linear-gradient(145deg,#131720,#10131a)!important;
  box-shadow:none!important;
  overflow:hidden;
}
details[data-testid="stExpander"] summary{
  color:#e9ecf3!important;
  font-weight:650!important;
}
details[data-testid="stExpander"] summary:hover{background:#181c26!important}

/* =========================
   TASK / ROADMAP
   ========================= */
.task-progress-shell,.roadmap-shell{
  background:
    radial-gradient(420px 130px at 95% 0%,rgba(139,124,255,.12),transparent 70%),
    linear-gradient(145deg,#151923,#10131a);
  border:1px solid #2a3040;
  border-radius:20px;
  padding:19px 21px;
  box-shadow:var(--shadow-sm);
  margin:10px 0 19px;
}
.task-progress-top{display:flex;justify-content:space-between;align-items:center;gap:12px;color:#e8ebf2}
.task-progress-bar,.roadmap-progress{height:9px;background:#252b37;border-radius:99px;overflow:hidden;margin-top:12px}
.task-progress-bar>div,.roadmap-progress>div{
  height:100%;
  background:linear-gradient(90deg,#8b7cff,#58d8d1);
  border-radius:99px;
  transition:width .35s ease;
  box-shadow:0 0 18px rgba(88,216,209,.18);
}
.roadmap-kicker{color:#a79bff;font-size:.7rem;text-transform:uppercase;letter-spacing:.14em;font-weight:800;margin-bottom:5px}
.roadmap-title{font-size:1.9rem;font-weight:800;color:#f2f4f8}
.roadmap-meta{color:#8992a7;margin-top:4px}
.roadmap-launch{
  background:
    radial-gradient(300px 100px at 90% 0%,rgba(139,124,255,.17),transparent 70%),
    linear-gradient(145deg,#171a28,#10131a);
  border:1px solid #353052;
  border-radius:20px;
  padding:18px 20px;
  display:flex;align-items:center;justify-content:space-between;gap:16px;
  margin:8px 0 18px;
  box-shadow:0 14px 35px rgba(0,0,0,.24);
}
.roadmap-launch-title{font-size:1.12rem;font-weight:800;margin-bottom:3px;color:#f0f2f7}
.roadmap-launch-sub{color:#8b94a8;font-size:.84rem}
.roadmap-week{
  margin:20px 0 16px;
  padding:16px 16px 9px;
  background:linear-gradient(145deg,#151923,#11141b);
  border:1px solid #292f3d;
  border-radius:17px;
  box-shadow:var(--shadow-sm);
}
.roadmap-week-title{font-size:1.18rem;font-weight:800;color:#eef1f7}
.roadmap-week-meta{color:#80899e;font-size:.78rem}
.roadmap-level{
  margin:11px 0;
  padding:11px 13px 7px;
  border:1px solid #282e3b;
  border-radius:13px;
  background:#11141b;
}
.roadmap-level-title{font-size:.95rem;font-weight:800;color:#dfe3ec}
.roadmap-topic-count{color:#778196;font-size:.77rem}
.roadmap-help{color:#858ea2;font-size:.84rem;margin:0 0 12px}
.completed-list{display:flex;flex-direction:column;gap:7px;margin-top:8px}
.completed-task{
  display:flex;align-items:center;gap:10px;padding:10px 12px;
  border-radius:12px;background:#11141a;border:1px solid #252b37;color:#747e91;
}
.completed-task .check{font-size:1rem;color:var(--mint);flex:none}
.completed-task .text{text-decoration:line-through;text-decoration-thickness:1.5px;flex:1}
.completed-task .meta{font-size:.72rem;color:#626c80;white-space:nowrap}

/* =========================
   REMINDERS / POMODORO
   ========================= */
.reminder-row{
  display:grid;grid-template-columns:32px 1fr auto;align-items:center;gap:10px;
  padding:11px 13px;border:1px solid #282e3b;
  background:linear-gradient(145deg,#151923,#11141b);
  border-radius:14px;margin-top:8px;
  transition:.18s ease;
}
.reminder-row:hover{border-color:#3a4050;transform:translateY(-1px)}
.reminder-text{font-size:.9rem;line-height:1.35;color:#e1e5ed}
.reminder-done .reminder-text{text-decoration:line-through;color:#687286}
.reminder-done .reminder-meta{color:var(--mint)}
.pomo-card{
  position:relative;
  overflow:hidden;
  background:
    radial-gradient(260px 160px at 50% 10%,rgba(255,113,133,.12),transparent 70%),
    radial-gradient(300px 180px at 50% 100%,rgba(139,124,255,.11),transparent 70%),
    linear-gradient(145deg,#17141e,#11131a);
  border:1px solid #342c3b;
  border-radius:22px;
  padding:20px;
  box-shadow:0 18px 50px rgba(0,0,0,.32),inset 0 1px rgba(255,255,255,.035);
}
.pomo-time{
  font-size:3.35rem;font-weight:800;letter-spacing:.035em;text-align:center;
  margin:12px 0 2px;color:#faf9fc;font-family:"DM Sans",Inter,sans-serif;
  text-shadow:0 0 30px rgba(255,113,133,.12);
}
.pomo-mode{text-align:center;color:#8e97aa;font-size:.72rem;text-transform:uppercase;letter-spacing:.13em}
.pomo-actions{display:flex;gap:7px;flex-wrap:wrap;justify-content:center;margin-top:14px}

/* =========================
   DATA EDITOR
   ========================= */
div[data-testid="stDataEditor"]{
  border:1px solid #2a3040!important;
  border-radius:16px!important;
  overflow:hidden;
  box-shadow:0 15px 40px rgba(0,0,0,.22)!important;
  background:#10131a!important;
}
div[data-testid="stDataEditor"] *{color:#dce1eb}
div[data-testid="stDataEditor"] [role="columnheader"]{
  background:#171b25!important;
  color:#aeb6c7!important;
}
div[data-testid="stDataEditor"] [role="gridcell"]{
  background:#10131a!important;
  border-color:#252b37!important;
}

/* =========================
   CHECKBOX / SLIDERS / PROGRESS
   ========================= */
.stCheckbox label{color:#dce1e9!important}
.stCheckbox input+div{border-color:#3b4353!important;background:#11141b!important}
.stCheckbox input:checked+div{background:#796aff!important;border-color:#796aff!important}
.stProgress>div>div>div{
  background:linear-gradient(90deg,#8b7cff,#50d7d4)!important;
  box-shadow:0 0 16px rgba(80,215,212,.15);
}

/* =========================
   LOGIN
   ========================= */
.login{
  max-width:450px;
  margin:11vh auto 0;
  text-align:center;
  padding:36px 34px;
  border:1px solid #302d4a;
  border-radius:26px;
  background:
    radial-gradient(300px 130px at 50% 0%,rgba(139,124,255,.15),transparent 75%),
    linear-gradient(145deg,#151824,#0f1219);
  box-shadow:0 30px 90px rgba(0,0,0,.48),inset 0 1px rgba(255,255,255,.04);
}
.login h1{font-size:2.45rem!important}
.login p{color:#8992a7}

/* =========================
   ALERTS / TOASTS
   ========================= */
div[data-testid="stAlert"]{
  border-radius:14px!important;
  border:1px solid #303746!important;
  background:#141820!important;
  color:#dfe4ed!important;
}
div[data-testid="stToast"]{
  background:#171b25!important;
  color:#edf0f5!important;
  border:1px solid #303746!important;
  box-shadow:0 18px 45px rgba(0,0,0,.4)!important;
}

/* =========================
   METRICS / TABS / TOOLBARS
   ========================= */
div[data-testid="stMetric"]{
  background:linear-gradient(145deg,#151923,#10131a);
  border:1px solid #272d3a;
  border-radius:16px;
  padding:13px 15px;
}
div[data-testid="stMetricLabel"]{color:#8791a5!important}
div[data-testid="stMetricValue"]{color:#f3f5f9!important}
button[data-baseweb="tab"]{
  color:#7f899d!important;
  font-weight:650!important;
}
button[data-baseweb="tab"][aria-selected="true"]{
  color:#bcb4ff!important;
}

/* =========================
   MOBILE
   ========================= */
@media(max-width:900px){
  .block-container{padding:1.25rem 1rem 3rem!important}
  .stMarkdown h1{font-size:1.8rem!important}
  .hero h1{font-size:1.9rem!important}
  .roadmap-title{font-size:1.55rem}
  .stat{min-height:94px}
}
@media(max-width:640px){
  section[data-testid="stSidebar"]{width:280px}
  .task-progress-top{align-items:flex-start;flex-direction:column}
  .completed-task .meta{display:none}
  .pomo-time{font-size:2.65rem}
  .hero{padding:22px 20px}
}
</style>
"""

st.markdown(CSS, unsafe_allow_html=True)


# =============================================================================
# AUTHENTICATION / WORKSPACE
# =============================================================================

# Passwords are stored as SHA-256 hashes rather than plaintext in the source.
# For a private Streamlit app this keeps the actual passwords out of the code.
AUTH_USERS = {
    "admin": {
        "display": "Admin",
        "password_hash": "d05a3ee9f244b6e64819cc6252a727193d80cf0fdfe0ac32db19f5baaca6fb75",
    },
    "kushal": {
        "display": "Kushal",
        "password_hash": "91e3a203cb36a25dc3b0ae1de7cb0aaf7ba8daced90c544400631e8310ad7afb",
    },
    "jahnavi": {
        "display": "Jahnavi",
        "password_hash": "030e4aa5b66a67e69eb979f11c744b5a839d96ee413f64f65a6e84caa79b9eb9",
    },
}

USER = None
NAME = "Tracker"


def _password_ok(role: str, password: str) -> bool:
    expected = AUTH_USERS[role]["password_hash"]
    actual = hashlib.sha256(password.encode("utf-8")).hexdigest()
    return hmac.compare_digest(actual, expected)


def login_gate() -> str:
    """Login with a maximum 24-hour session; normal Streamlit reruns do not log out."""
    current_user = st.session_state.get("authenticated_user")
    login_at = st.session_state.get("login_at")
    if current_user in AUTH_USERS and login_at:
        try:
            if (dt.datetime.now() - dt.datetime.fromisoformat(str(login_at))).total_seconds() < 86400:
                return current_user
        except (TypeError, ValueError):
            pass
        st.session_state.pop("authenticated_user", None)
        st.session_state.pop("login_at", None)
        st.session_state.pop("db", None)
        st.session_state.pop("_supabase_snapshot", None)

    st.markdown('<div class="login"><h1>✅ Tracker</h1><p>Sign in to continue</p></div>', unsafe_allow_html=True)
    with st.form("login_form", clear_on_submit=False):
        role = st.selectbox("Role", list(AUTH_USERS.keys()), format_func=lambda x: AUTH_USERS[x]["display"])
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Login", type="primary", width="stretch")

        if submitted:
            if _password_ok(role, password):
                st.session_state["authenticated_user"] = role
                st.session_state["login_at"] = dt.datetime.now().isoformat(timespec="seconds")
                st.session_state.pop("db", None)
                st.session_state.pop("_supabase_snapshot", None)
                st.rerun()
            else:
                st.error("Incorrect role or password.")

    st.stop()


USER = login_gate()


# =============================================================================
# STORAGE — SUPABASE
# =============================================================================

LOCK = threading.RLock()

# Supabase table names. Each table stores the app's flexible record payload in
# JSONB so the existing Streamlit editor can keep its current column names
# (including names containing spaces) without a large UI rewrite.
SUPABASE_TABLES = {
    "Tasks": "tasks",
    "Workout": "workout",
    "Wishlist": "wishlist",
    "Jobs": "jobs",
    "Ideas": "ideas",
    "Projects": "projects",
    "Random": "random_notes",
    "Roadmap": "roadmap",
    "Habits": "habits",
    "HabitLog": "habit_log",
    "Targets": "targets",
    "Reminders": "reminders",
    "WishNotes": "wish_notes",
}


@st.cache_resource
def supabase_client():
    """Create one reusable Supabase client per Streamlit process."""
    if create_client is None:
        raise RuntimeError(
            "The 'supabase' package is missing. Add 'supabase' to requirements.txt."
        )

    if "supabase" not in st.secrets:
        raise RuntimeError(
            "Missing [supabase] in Streamlit secrets. "
            "Set url and key in .streamlit/secrets.toml."
        )

    config = st.secrets["supabase"]
    url = str(config.get("url", "")).strip()
    key = str(config.get("key", "")).strip()

    if not url or not key:
        raise RuntimeError(
            "Supabase secrets must contain both 'url' and 'key'."
        )

    return create_client(url, key)


def _json_safe(value: Any) -> Any:
    """Convert pandas/NumPy/Python values into strict JSON-safe values.

    Supabase JSONB rejects IEEE-754 NaN/Infinity values. Streamlit's
    data_editor can produce those values for empty numeric cells, so every
    value is sanitized immediately before it reaches the API.
    """
    if value is None:
        return None
    if isinstance(value, float):
        return value if math.isfinite(value) else None
    if isinstance(value, (dt.datetime, dt.date)):
        return value.isoformat()
    if isinstance(value, dt.time):
        return value.strftime("%H:%M:%S")
    if isinstance(value, dict):
        return {str(k): _json_safe(v) for k, v in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [_json_safe(v) for v in value]
    # pandas / NumPy scalar values (including numpy.nan).
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    if hasattr(value, "item"):
        try:
            return _json_safe(value.item())
        except (ValueError, TypeError):
            pass
    return value


def normalize_record(sheet: str, record: dict[str, Any]) -> dict[str, Any]:
    """Normalize a record to the app schema before storing or displaying it."""
    schema = SHEETS[sheet]
    out: dict[str, Any] = {}

    for col, (kind, _) in schema.items():
        value = record.get(col)
        if value is None:
            out[col] = ""
        elif kind == "num":
            try:
                n = float(value)
                # Empty numeric cells can arrive from pandas as NaN.
                if not math.isfinite(n):
                    out[col] = ""
                else:
                    out[col] = int(n) if n.is_integer() else n
            except (ValueError, TypeError):
                out[col] = ""
        elif kind == "date":
            try:
                if pd.isna(value):
                    out[col] = ""
                elif isinstance(value, dt.datetime):
                    out[col] = value.date().isoformat()
                elif isinstance(value, dt.date):
                    out[col] = value.isoformat()
                else:
                    out[col] = str(value)
            except (TypeError, ValueError):
                out[col] = ""
        elif kind == "time":
            if isinstance(value, dt.time):
                out[col] = value.strftime("%H:%M")
            else:
                out[col] = "" if str(value).lower() in {"nan", "nat", "none"} else str(value)
        else:
            text_value = str(value)
            out[col] = "" if text_value.lower() in {"nan", "nat", "none"} else text_value

    out["_id"] = str(record.get("_id") or uuid.uuid4())
    if record.get("__owner"):
        out["__owner"] = str(record["__owner"])
    if record.get("__created"):
        out["__created"] = str(record["__created"])
    return _json_safe(out)


def _blank(v: Any) -> bool:
    """True for None / NaN / NaT / empty or whitespace-only strings."""
    if v is None:
        return True
    try:
        if pd.isna(v):
            return True
    except (TypeError, ValueError):
        pass
    return str(v).strip() == ""


def _public(rec: dict[str, Any]) -> dict[str, Any]:
    """Record without internal bookkeeping keys (__owner, __created)."""
    return {k: v for k, v in rec.items() if not str(k).startswith("__")}


def _record_is_blank(sheet: str, rec: dict[str, Any]) -> bool:
    return all(_blank(rec.get(c)) for c in SHEETS[sheet])


def _can_modify(rec: dict[str, Any]) -> bool:
    """Admin can modify everything; users only their own private records."""
    return USER == "admin" or str(rec.get("__owner") or USER) == USER


def validate_record(sheet: str, record: dict[str, Any]) -> tuple[bool, str]:
    """Validate a record against the app schema before saving."""
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
                s = dt.datetime.strptime(str(start), "%H:%M").time()
                e = dt.datetime.strptime(str(end), "%H:%M").time()
                if e <= s:
                    return False, "Tasks: End time must be after Start time"
            except ValueError:
                return False, "Tasks: invalid time"

    if sheet == "Jobs":
        link = str(record.get("Link", "")).strip()
        if link and not re.match(r"^https?://", link, re.I):
            return False, "Jobs: Link must start with http:// or https://"

    return True, ""


def _remote_row(record: dict[str, Any], owner: str | None = None) -> dict[str, Any]:
    """Convert an app record into the compact Supabase row."""
    normalized = dict(record)
    record_id = str(normalized.pop("_id", "") or uuid.uuid4())
    record_owner = str(owner or normalized.pop("__owner", "") or USER or "admin")
    normalized.pop("__owner", None)
    created = normalized.pop("__created", None)
    # Final defensive sanitization: Supabase/PostgREST JSON encoding does
    # not accept NaN or +/-Infinity, even though pandas/NumPy can create them.
    normalized = _json_safe(normalized)
    if created:
        # Creation time keeps rows in insertion order after reloads/updates.
        normalized["_created"] = str(created)
    return {
        "_id": record_id,
        "owner": record_owner,
        "data": normalized,
    }


def _decode_remote_row(sheet: str, row: dict[str, Any]) -> dict[str, Any]:
    """Convert a Supabase JSONB row back into the app's record format."""
    data = row.get("data") or {}
    if not isinstance(data, dict):
        data = {}
    record = dict(data)
    record["_id"] = str(row.get("_id") or uuid.uuid4())
    record["__owner"] = str(row.get("owner") or "admin")
    if data.get("_created"):
        record["__created"] = str(data["_created"])
    return normalize_record(sheet, record)


# Progress/logs are private to each user. Definitions (Habits, Targets, ...) stay
# admin-shared so every account sees the same habit list and workout targets.
PER_USER_SHEETS = {"HabitLog", "Reminders", "WishNotes", "Roadmap"}


def _fetch_all(client, table: str, owners: list[str]) -> list[dict[str, Any]]:
    """Fetch every row for the given owners.

    PostgREST returns at most 1000 rows per request by default, which silently
    truncated large tables (the roadmap has ~800 rows per user). The first page is
    fetched in natural order; only if it is full do we re-fetch with a stable
    ordering and page through the rest.
    """
    page = 1000

    def query():
        return client.table(table).select("_id,owner,data").in_("owner", owners)

    first = query().range(0, page - 1).execute().data or []
    if len(first) < page:
        return first

    rows: list[dict[str, Any]] = []
    start = 0
    while True:
        batch = query().order("_id").range(start, start + page - 1).execute().data or []
        rows.extend(batch)
        if len(batch) < page:
            break
        start += page
    return rows


def supabase_load(user: str | None = None) -> dict[str, list[dict[str, Any]]]:
    """Load shared admin records plus the current user's private records."""
    user = user or USER or "admin"
    client = supabase_client()
    db = {sheet: [] for sheet in SHEETS}
    snapshot = {}

    for sheet, table in SUPABASE_TABLES.items():
        owners = [user] if sheet in PER_USER_SHEETS else sorted({"admin", user})
        rows = _fetch_all(client, table, owners)
        decoded = [_decode_remote_row(sheet, row) for row in rows]
        # Stable sort: rows created by this version keep insertion order;
        # older rows (no timestamp) keep whatever order the database returned.
        decoded.sort(key=lambda r: r.get("__created") or "")
        db[sheet] = decoded
        snapshot[sheet] = {str(r["_id"]): dict(r) for r in db[sheet]}

    st.session_state["_supabase_snapshot"] = snapshot
    return db


def supabase_save(
    user: str | None = None,
    db: dict[str, list[dict[str, Any]]] | None = None,
) -> None:
    """Persist changes while enforcing admin-shared/user-private ownership."""
    user = user or USER or "admin"
    if db is None:
        db = {sheet: [] for sheet in SHEETS}

    client = supabase_client()
    previous = st.session_state.get("_supabase_snapshot", {})
    blocked = 0

    for sheet, table in SUPABASE_TABLES.items():
        old_records = previous.get(sheet, {})
        current_records = {}

        for raw in db.get(sheet, []):
            record = normalize_record(sheet, raw)
            rid = str(record.get("_id") or uuid.uuid4())
            old = old_records.get(rid)
            owner = str(
                record.get("__owner")
                or (old or {}).get("__owner")
                or user
            )
            # A non-admin can only create/change their own private records.
            # Existing admin records remain shared and immutable to them.
            if user != "admin" and old and str(old.get("__owner")) == "admin":
                if _public(record) != _public(old):
                    blocked += 1
                record = dict(old)
                owner = "admin"
            if old is None and not record.get("__created"):
                record["__created"] = dt.datetime.now().isoformat(timespec="microseconds")
            record["__owner"] = owner
            record["_id"] = rid
            current_records[rid] = record

        # Restore any shared admin record that a normal user tried to delete.
        if user != "admin":
            for rid, old in old_records.items():
                if str(old.get("__owner")) == "admin" and rid not in current_records:
                    current_records[rid] = dict(old)
                    blocked += 1

        deleted_ids = []
        for rid, old in old_records.items():
            if rid not in current_records:
                if user == "admin" or str(old.get("__owner")) == user:
                    deleted_ids.append(rid)

        if deleted_ids:
            for i in range(0, len(deleted_ids), 100):
                client.table(table).delete().in_("_id", deleted_ids[i:i + 100]).execute()

        changed = []
        for rid, record in current_records.items():
            old = old_records.get(rid)
            owner = str(record.get("__owner") or user)
            # Only admin can modify shared admin-owned records.
            if user != "admin" and owner == "admin":
                continue
            if old != record:
                changed.append(_remote_row(record, owner=owner))

        if changed:
            for i in range(0, len(changed), 100):
                client.table(table).upsert(
                    changed[i:i + 100], on_conflict="_id"
                ).execute()

        previous[sheet] = {rid: dict(record) for rid, record in current_records.items()}
        db[sheet] = list(current_records.values())

    st.session_state["_supabase_snapshot"] = previous
    if blocked:
        st.session_state["_blocked_notice"] = True


def seed_if_empty(db: dict[str, list[dict[str, Any]]]) -> bool:
    changed = False
    for sheet, defaults in DEFAULTS.items():
        if not db[sheet]:
            db[sheet] = [normalize_record(sheet, dict(r)) for r in defaults]
            changed = True
    return changed


def _purge_blank_records(db: dict[str, list[dict[str, Any]]]) -> bool:
    """Drop all-blank rows left behind by the old 'save on +' editor bug.

    Roadmap is skipped on purpose: its un-converted JSON section rows decode to
    blank records and must survive until roadmap_convert.sql has been run.
    """
    changed = False
    for sheet, rows in db.items():
        if sheet == "Roadmap":
            continue
        keep = [
            r for r in rows
            if not (_record_is_blank(sheet, r) and _can_modify(r))
        ]
        if len(keep) != len(rows):
            db[sheet] = keep
            changed = True
    return changed


def load_db() -> dict[str, list[dict[str, Any]]]:
    key = "db"
    if key in st.session_state:
        return st.session_state[key]

    with LOCK:
        try:
            db = supabase_load(USER)
        except Exception as exc:
            st.error(
                "Could not connect to Supabase. "
                "Check your Streamlit secrets and database tables."
            )
            st.exception(exc)
            st.stop()

    changed = _purge_blank_records(db)

    # Seed the default starter data only from the admin account so defaults
    # become shared records instead of accidentally becoming private to a user.
    if USER == "admin" and seed_if_empty(db):
        changed = True

    if changed:
        try:
            supabase_save(USER, db)
        except Exception as exc:
            st.error("Supabase connected, but cleanup/initial data could not be saved.")
            st.exception(exc)
            st.stop()

    st.session_state[key] = db
    return db


def save_db(db: dict[str, list[dict[str, Any]]]) -> None:
    with LOCK:
        try:
            supabase_save(USER, db)
        except Exception as exc:
            st.error("Could not save changes to Supabase.")
            st.exception(exc)
            st.stop()
    # Lets the sidebar rebuild the Excel backup only when data actually changed.
    st.session_state["_db_version"] = st.session_state.get("_db_version", 0) + 1


DB = load_db()


# =============================================================================
# SERIALIZATION / EDITOR
# =============================================================================

def to_editor_df(sheet: str, rows: list[dict[str, Any]]) -> pd.DataFrame:
    """Rows -> DataFrame for st.data_editor. Carries a hidden `_id` column so
    edited rows are matched to records by identity, never by position."""
    schema = SHEETS[sheet]
    data = []

    for row in rows:
        r = normalize_record(sheet, row)
        item = {"_id": r["_id"]}
        item.update({c: r.get(c, "") for c in schema})

        for c, (kind, _) in schema.items():
            if kind == "date":
                try:
                    item[c] = pd.to_datetime(item[c]).date() if item[c] else None
                except Exception:
                    item[c] = None
            elif kind == "time":
                try:
                    # accepts "HH:MM" and "HH:MM:SS"
                    item[c] = dt.datetime.strptime(str(item[c])[:5], "%H:%M").time() if item[c] else None
                except Exception:
                    item[c] = None
            elif kind == "num":
                try:
                    item[c] = None if item[c] in ("", None) else float(item[c])
                except Exception:
                    item[c] = None

        data.append(item)

    # Always reset the index before passing data to st.data_editor.
    return pd.DataFrame(data, columns=["_id"] + list(schema)).reset_index(drop=True)


def sort_tasks_df(df: pd.DataFrame) -> pd.DataFrame:
    """Order tasks by date, then start time. Undated / untimed tasks go last."""
    if df.empty:
        return df.reset_index(drop=True)
    d = df.copy()

    def date_key(x):
        return x.isoformat() if isinstance(x, dt.date) and not _blank(x) else "9999-99-99"

    def time_key(x):
        if _blank(x):
            return "99:99"
        return x.strftime("%H:%M") if isinstance(x, dt.time) else str(x)[:5]

    d["_kd"] = d["Date"].apply(date_key)
    d["_kt"] = d["Start"].apply(time_key)
    return (
        d.sort_values(["_kd", "_kt"], kind="stable")
        .drop(columns=["_kd", "_kt"])
        .reset_index(drop=True)
    )


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


def editor(
    sheet: str,
    mask=None,
    defaults: dict[str, Any] | None = None,
    show: list[str] | None = None,
    key: str | None = None,
) -> None:
    """Editable table backed by DB[sheet].

    Fixes over the previous version:
      * rows are matched by a hidden `_id`, not by position, so deleting row 1
        deletes row 1 (and not whatever happens to shift into its place);
      * the blank row created by clicking "+" is never saved, so typing into it
        no longer needs to be done twice;
      * after a real save the widget key is bumped so the editor re-baselines
        from the database (and re-sorts Tasks by time).
    """
    schema = SHEETS[sheet]
    cols = show or list(schema)
    defaults = dict(defaults or {})
    if sheet == "Tasks":
        defaults.setdefault("Status", "To do")

    base_key = key or f"editor_{sheet}_{'_'.join(cols)}"
    ver_key = f"{base_key}__ver"
    ver = st.session_state.get(ver_key, 0)

    full_df = to_editor_df(sheet, DB[sheet])
    if full_df.empty:
        keep = pd.Series([], dtype=bool)
    else:
        # Never show all-blank leftovers from the old bug.
        is_blank = full_df[list(schema)].apply(lambda col: col.map(_blank)).all(axis=1)
        keep = ~is_blank
        if mask is not None:
            keep = keep & mask(full_df).astype(bool)

    visible = full_df[keep]
    if sheet == "Tasks":
        visible = sort_tasks_df(visible)
    visible = visible.reset_index(drop=True)

    edited = st.data_editor(
        visible[["_id"] + cols],
        num_rows="dynamic",
        width="stretch",
        hide_index=True,
        key=f"{base_key}_{ver}",
        disabled=["_id"],
        column_config={"_id": None, **column_config_for(sheet, cols)},
    )

    old_by_id = {str(r.get("_id")): r for r in DB[sheet]}
    visible_ids = set(visible["_id"].astype(str))
    new_visible: list[dict[str, Any]] = []

    for _, row in edited.reset_index(drop=True).iterrows():
        rid = None if _blank(row.get("_id")) else str(row["_id"])
        old = old_by_id.get(rid) if rid else None

        rec = {c: row.get(c) for c in cols}

        # Skip the empty row that appears right after clicking "+".
        if all(_blank(rec.get(c)) for c in cols if c not in defaults):
            continue

        merged = dict(old) if old else {}
        merged.update(rec)
        for c, d in defaults.items():          # includes hidden columns (e.g. Workout.Target)
            if _blank(merged.get(c)):
                merged[c] = d
        merged["_id"] = rid or str(uuid.uuid4())

        if sheet == "Tasks":
            old_status = str((old or {}).get("Status", ""))
            if str(merged.get("Status", "")) == "Done":
                if old_status != "Done" and _blank(merged.get("CompletedAt")):
                    merged["CompletedAt"] = dt.datetime.now().isoformat(timespec="seconds")
            else:
                merged["CompletedAt"] = ""

        merged = normalize_record(sheet, merged)

        # Only validate rows the user actually changed, so one legacy bad row
        # cannot block every edit in the table.
        old_norm = normalize_record(sheet, old) if old else None
        if old_norm is None or _public(merged) != _public(old_norm):
            ok, message = validate_record(sheet, merged)
            if not ok:
                st.error(message)
                return

        new_visible.append(merged)

    hidden = [r for r in DB[sheet] if str(r.get("_id")) not in visible_ids]
    combined = hidden + new_visible

    def as_map(rows):
        return {str(r["_id"]): r for r in serialize_records(sheet, rows)}

    if as_map(DB[sheet]) != as_map(combined):
        DB[sheet] = combined
        save_db(DB)
        st.session_state[ver_key] = ver + 1
        st.rerun()


def serialize_records(sheet: str, rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [
        normalize_record(sheet, r)
        for r in rows
    ]


# =============================================================================
# TASKS / QUICK ADD
# =============================================================================

def cleanup_completed_tasks() -> None:
    """Delete completed tasks seven days after their completion timestamp."""
    changed = False
    cutoff = dt.datetime.now() - dt.timedelta(days=7)
    kept = []
    for r in DB["Tasks"]:
        if str(r.get("Status", "")) == "Done" and _can_modify(r):
            stamp = str(r.get("CompletedAt", "")).strip()
            if stamp:
                try:
                    if dt.datetime.fromisoformat(stamp) < cutoff:
                        changed = True
                        continue
                except ValueError:
                    pass
        kept.append(r)
    if changed:
        DB["Tasks"] = kept
        save_db(DB)


def task_completion_progress() -> tuple[int, int, int]:
    cleanup_completed_tasks()
    total = len(DB["Tasks"])
    done = sum(str(r.get("Status", "")) == "Done" for r in DB["Tasks"])
    pct = round(100 * done / total) if total else 0
    return done, total, pct


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
                    # keep CompletedAt in sync so the 7-day cleanup works
                    if new_status == "Done":
                        if not str(r.get("CompletedAt", "")).strip():
                            r["CompletedAt"] = dt.datetime.now().isoformat(timespec="seconds")
                    else:
                        r["CompletedAt"] = ""
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
        today_all = today_df = overdue = df
    else:
        today_all = df[df["Date"] == TODAY]
        today_df = sort_tasks_df(today_all[today_all["Status"] != "Done"])
        overdue = sort_tasks_df(
            df[
                df["Date"].notna()
                & (df["Date"] < TODAY)
                & (df["Status"] != "Done")
            ]
        )

    work = df[df["Area"].isin(["DSA", "Python", "Analytics", "Project", "Work"])] if not df.empty else df
    # Done tasks are filtered out of today_df, so count them from today_all.
    done_today = len(today_all) - len(today_df)
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
        stat("Tasks today", f"{done_today}/{len(today_all)}",
             pct=round(100 * done_today / len(today_all)) if len(today_all) else 0),
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


def _toggle_habit(day: dt.date, habit: str, widget_key: str) -> None:
    """Checkbox callback: read the widget's real value instead of stale args."""
    set_habit(day, habit, bool(st.session_state.get(widget_key)))


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
                on_change=_toggle_habit,
                args=(TODAY, habit, key),
            )


def _cleanup_reminders() -> None:
    """Delete reminders only after the calendar day on which they were completed."""
    changed = False
    today_key = TODAY.isoformat()
    kept = []
    for r in DB["Reminders"]:
        done_at = str(r.get("DoneAt", "")).strip()
        if done_at and _can_modify(r):
            try:
                done_day = dt.datetime.fromisoformat(done_at).date().isoformat()
                if done_day < today_key:
                    changed = True
                    continue
            except ValueError:
                pass
        kept.append(r)
    if changed:
        DB["Reminders"] = kept
        save_db(DB)


def reminders() -> None:
    _cleanup_reminders()
    with st.container(border=True):
        st.markdown("### 📝 Reminders")
        st.caption("Checked reminders stay visible until 12:00 AM, then clear automatically.")
        with st.form("reminder_form", clear_on_submit=True):
            c1, c2 = st.columns([5, 1])
            value = c1.text_input("Reminder", label_visibility="collapsed", placeholder="Add a reminder…")
            add = c2.form_submit_button("＋ Add", type="primary", width="stretch")
            if add and value.strip():
                DB["Reminders"].append({
                    "_id": str(uuid.uuid4()), "Text": value.strip(),
                    "Done": "", "DoneAt": "",
                })
                save_db(DB)
                st.rerun()

        for r in DB["Reminders"]:
            rid = str(r["_id"])
            done = bool(str(r.get("DoneAt", "")).strip())
            c1, c2, c3 = st.columns([0.7, 8.0, 1.2], vertical_alignment="center")
            checked = c1.checkbox("", value=done, key=f"rem_check_{rid}", label_visibility="collapsed")
            if checked != done:
                if checked:
                    r["Done"] = "1"
                    r["DoneAt"] = dt.datetime.now().isoformat(timespec="seconds")
                else:
                    r["Done"] = ""
                    r["DoneAt"] = ""
                save_db(DB)
                st.rerun()
            c2.markdown(
                f'<div class="reminder-text {"reminder-done" if done else ""}">{esc(str(r.get("Text", "")))}'
                f'<div class="reminder-meta">{"✓ checked — clears at midnight" if done else "today"}</div></div>',
                unsafe_allow_html=True,
            )
            if c3.button("✕", key=f"rem_delete_{rid}"):
                DB["Reminders"] = [x for x in DB["Reminders"] if x["_id"] != rid]
                save_db(DB)
                st.rerun()


def pomodoro() -> None:
    """Reliable session-state Pomodoro with focus, break, rest and reset controls."""
    defaults = {
        "pomo_total": 25 * 60, "pomo_remaining": 25 * 60,
        "pomo_mode": "Focus", "pomo_running": False, "pomo_end_at": None,
    }
    for k, v in defaults.items():
        st.session_state.setdefault(k, v)

    @st.fragment(run_every="1s")
    def _timer():
        now = time.time()
        if st.session_state.pomo_running and st.session_state.pomo_end_at:
            remaining = max(0, int(st.session_state.pomo_end_at - now))
            st.session_state.pomo_remaining = remaining
            if remaining <= 0:
                st.session_state.pomo_running = False
                st.session_state.pomo_end_at = None
                st.session_state.pomo_remaining = 0

        total = max(1, int(st.session_state.pomo_total))
        remaining = int(st.session_state.pomo_remaining)
        mins, secs = divmod(remaining, 60)
        pct = max(0, min(100, round(100 * (1 - remaining / total))))
        st.markdown(
            f'<div class="pomo-card"><div style="font-weight:700">🍅 Pomodoro</div>'
            f'<div class="pomo-time">{mins:02d}:{secs:02d}</div>'
            f'<div class="pomo-mode">{esc(st.session_state.pomo_mode)} · {pct}%</div>'
            f'<div class="roadmap-progress" style="margin-top:12px"><div style="width:{pct}%"></div></div></div>',
            unsafe_allow_html=True,
        )
        a, b, c, d = st.columns(4)
        if a.button("▶ Start" if not st.session_state.pomo_running else "⏸ Pause", key="pomo_start", width="stretch"):
            if st.session_state.pomo_running:
                st.session_state.pomo_running = False
                st.session_state.pomo_end_at = None
            else:
                st.session_state.pomo_end_at = time.time() + st.session_state.pomo_remaining
                st.session_state.pomo_running = True
            st.rerun()
        if b.button("☕ Rest", key="pomo_rest", width="stretch"):
            st.session_state.pomo_total = 5 * 60
            st.session_state.pomo_remaining = 5 * 60
            st.session_state.pomo_mode = "Rest"
            st.session_state.pomo_running = False
            st.session_state.pomo_end_at = None
            st.rerun()
        if c.button("25 Focus", key="pomo_focus", width="stretch"):
            st.session_state.pomo_total = 25 * 60
            st.session_state.pomo_remaining = 25 * 60
            st.session_state.pomo_mode = "Focus"
            st.session_state.pomo_running = False
            st.session_state.pomo_end_at = None
            st.rerun()
        if d.button("↺ Reset", key="pomo_reset", width="stretch"):
            st.session_state.pomo_total = 25 * 60
            st.session_state.pomo_remaining = 25 * 60
            st.session_state.pomo_mode = "Focus"
            st.session_state.pomo_running = False
            st.session_state.pomo_end_at = None
            st.rerun()

        x, y = st.columns(2)
        if x.button("5 min break", key="pomo_5", width="stretch"):
            st.session_state.pomo_total = 5 * 60
            st.session_state.pomo_remaining = 5 * 60
            st.session_state.pomo_mode = "Break"
            st.session_state.pomo_running = False
            st.session_state.pomo_end_at = None
            st.rerun()
        if y.button("15 min break", key="pomo_15", width="stretch"):
            st.session_state.pomo_total = 15 * 60
            st.session_state.pomo_remaining = 15 * 60
            st.session_state.pomo_mode = "Long break"
            st.session_state.pomo_running = False
            st.session_state.pomo_end_at = None
            st.rerun()

    _timer()


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

    # The key changes after every save (and when the habit list changes) so the
    # grid always re-baselines from the database instead of replaying stale edits.
    names_key = hashlib.sha1("|".join(names).encode()).hexdigest()[:6]
    grid_ver = st.session_state.get("habits_grid_v", 0)

    ed = st.data_editor(
        pd.DataFrame(rows),
        hide_index=True,
        width="stretch",
        disabled=["Day", "Done %"],
        key=f"habits_grid_{grid_ver}_{names_key}",
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
        st.session_state["habits_grid_v"] = grid_ver + 1
        st.rerun()

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
        editor("Habits", key="habits_settings")


def tasks_page() -> None:
    cleanup_completed_tasks()
    st.title("📋 Master Tasklist")

    done, total, pct = task_completion_progress()
    active = total - done
    st.markdown(
        f'<div class="task-progress-shell"><div class="task-progress-top">'
        f'<b>Task progress</b><span style="color:#8b90a3">{done}/{total} completed · {pct}%</span></div>'
        f'<div class="task-progress-bar"><div style="width:{pct}%"></div></div>'
        f'<div style="color:#70768a;font-size:.76rem;margin-top:7px">{active} active · completed tasks are archived for 7 days</div></div>',
        unsafe_allow_html=True,
    )

    with st.expander("➕ Quick add", expanded=False):
        quick_add(TODAY, "tasks")

    areas = st.multiselect("Filter active tasks by area", AREAS, key="task_filter")
    def active_mask(d):
        keep = d["Status"].astype(str) != "Done"
        if areas:
            keep = keep & d["Area"].isin(areas)
        return keep

    st.markdown("### Active tasks")
    editor(
        "Tasks", mask=active_mask, key=f"tasks_editor_{'_'.join(areas)}",
        show=["Task", "Area", "Status", "Date", "Start", "End", "Deadline", "Minutes"],
    )

    completed = [r for r in DB["Tasks"] if str(r.get("Status", "")) == "Done"]
    if completed:
        st.markdown("### ✓ Completed master list")
        st.caption("Completed tasks stay crossed out here for 7 days before automatic cleanup.")
        rows = []
        for r in sorted(completed, key=lambda x: str(x.get("CompletedAt", "")), reverse=True):
            stamp = str(r.get("CompletedAt", ""))
            meta = "Completed"
            if stamp:
                try:
                    when = dt.datetime.fromisoformat(stamp)
                    age = max(0, (dt.datetime.now() - when).days)
                    meta = f"Completed {when:%d %b} · {max(0, 7-age)}d left"
                except ValueError:
                    pass
            rows.append(
                f'<div class="completed-task"><span class="check">✓</span>'
                f'<span class="text">{esc(str(r.get("Task", "")))}</span>'
                f'<span class="meta">{esc(str(r.get("Area", "")))} · {esc(meta)}</span></div>'
            )
        st.markdown('<div class="completed-list">'+''.join(rows)+'</div>', unsafe_allow_html=True)


@st.dialog("🗺️ Roadmap", width="large")
def roadmap_dialog(rid: str, title: str) -> None:
    """Notion-style checklist roadmap.

    Roadmap records use:
      Roadmap = roadmap id          Heading = section / week
      Sub     = level (Easy, ...)   Topic   = checklist item
      Done    = "1" when completed  Order   = display order

    Every rerun in here is scoped to the dialog fragment, so ticking a box or
    adding a topic no longer closes the dialog.
    """

    def _order(r: dict[str, Any]) -> float:
        try:
            return float(r.get("Order") or 1e12)
        except (TypeError, ValueError):
            return 1e12

    def _refresh() -> None:
        st.rerun(scope="fragment")

    rows = sorted(
        (r for r in DB["Roadmap"] if r.get("Roadmap") == rid),
        key=_order,
    )
    items = [r for r in rows if str(r.get("Topic", "")).strip()]

    done = sum(str(r.get("Done", "")) == "1" for r in items)
    total = len(items)
    pct = round(100 * done / total) if total else 0

    headings = list(dict.fromkeys(
        str(r.get("Heading", "")).strip()
        for r in rows
        if str(r.get("Heading", "")).strip()
    ))

    st.markdown(
        f"""
        <div class="roadmap-shell">
            <div class="roadmap-kicker">Learning roadmap</div>
            <div class="roadmap-title">🗺️ {esc(title)}</div>
            <div class="roadmap-meta">{done} of {total} topics completed · {pct}% complete</div>
            <div class="roadmap-progress"><div style="width:{pct}%"></div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if not headings:
        st.info(
            "This roadmap is empty. Use **Roadmap settings** below to create a section, "
            "or run roadmap_convert.sql in Supabase to load the DSA roadmap."
        )
    else:
        # One section at a time keeps the dialog fast (the DSA roadmap alone has
        # ~800 checkboxes). Widget labels stay constant so open/closed state and
        # the selection survive a tick.
        sec_key = f"roadmap_section_{rid}"
        if st.session_state.get(sec_key) not in headings:
            st.session_state.pop(sec_key, None)
        heading = st.selectbox("Section", headings, key=sec_key)

        heading_items = [
            r for r in items if str(r.get("Heading", "")).strip() == heading
        ]
        h_done = sum(str(r.get("Done", "")) == "1" for r in heading_items)
        h_total = len(heading_items)
        h_pct = round(100 * h_done / h_total) if h_total else 0
        st.markdown(
            f'<div class="roadmap-week"><div class="roadmap-week-title">{esc(heading)}</div>'
            f'<div class="roadmap-week-meta">{h_done}/{h_total} topics · {h_pct}%</div>'
            f'<div class="roadmap-progress"><div style="width:{h_pct}%"></div></div></div>',
            unsafe_allow_html=True,
        )

        subs = list(dict.fromkeys(
            str(r.get("Sub", "")).strip() for r in heading_items
        ))
        # topics without a level first, then levels in their stored order
        subs.sort(key=lambda s: s != "")

        for sub in subs:
            sub_items = [
                r for r in heading_items if str(r.get("Sub", "")).strip() == sub
            ]
            sub_done = sum(str(r.get("Done", "")) == "1" for r in sub_items)
            with st.expander(sub or "Topics", expanded=False):
                st.caption(f"{sub_done}/{len(sub_items)} completed")
                for r in sub_items:
                    topic_id = str(r.get("_id"))
                    checked = str(r.get("Done", "")) == "1"
                    value = st.checkbox(
                        str(r.get("Topic", "")),
                        value=checked,
                        key=f"roadmap_topic_{rid}_{topic_id}",
                    )
                    if value != checked:
                        r["Done"] = "1" if value else ""
                        save_db(DB)
                        _refresh()

    with st.expander("⚙️ Roadmap settings"):
        st.markdown(
            '<p class="roadmap-help">'
            'Build the same hierarchy as Notion: <b>Week/Section → Level → Topics</b>. '
            'Put one topic per line.'
            '</p>',
            unsafe_allow_html=True,
        )

        used_orders = [_order(r) for r in rows if _order(r) < 1e12]
        next_order = (max(used_orders) if used_orders else 0) + 1

        # ---- Add a new week/section ----
        st.markdown("#### Add week / section")
        new_heading = st.text_input(
            "Section name",
            placeholder="e.g. Week 1: Arrays",
            key=f"roadmap_new_heading_{rid}",
        )
        if st.button("＋ Add section", key=f"roadmap_add_heading_{rid}", type="primary"):
            value = new_heading.strip()
            if not value:
                st.warning("Enter a section name first.")
            elif value in headings:
                st.warning("That section already exists.")
            else:
                DB["Roadmap"].append(normalize_record("Roadmap", {
                    "Roadmap": rid, "Heading": value, "Sub": "",
                    "Topic": "", "Done": "", "Order": next_order,
                }))
                save_db(DB)
                _refresh()

        if headings:
            st.divider()

            # ---- Add a difficulty/sub-section ----
            st.markdown("#### Add level / sub-section")
            selected_heading = st.selectbox(
                "Section", headings, key=f"roadmap_manage_heading_{rid}",
            )
            existing_subs = list(dict.fromkeys(
                str(r.get("Sub", "")).strip()
                for r in rows
                if str(r.get("Heading", "")).strip() == selected_heading
                and str(r.get("Sub", "")).strip()
            ))
            new_sub = st.text_input(
                "Level name", placeholder="Easy / Medium / Hard",
                key=f"roadmap_new_sub_{rid}",
            )
            if st.button("＋ Add level", key=f"roadmap_add_sub_{rid}"):
                value = new_sub.strip()
                if not value:
                    st.warning("Enter a level name first.")
                elif value in existing_subs:
                    st.warning("That level already exists in this section.")
                else:
                    DB["Roadmap"].append(normalize_record("Roadmap", {
                        "Roadmap": rid, "Heading": selected_heading, "Sub": value,
                        "Topic": "", "Done": "", "Order": next_order,
                    }))
                    save_db(DB)
                    _refresh()

            # ---- Add checklist topics ----
            st.divider()
            st.markdown("#### Add checklist topics")
            selected_sub = st.selectbox(
                "Level", ["(none)"] + existing_subs, key=f"roadmap_manage_sub_{rid}",
            )
            topics = st.text_area(
                "Topics — one per line",
                placeholder="Two Sum\nBest Time to Buy and Sell Stock\nContains Duplicate",
                height=150,
                key=f"roadmap_topics_{rid}",
            )
            if st.button("＋ Add topics", key=f"roadmap_add_topics_{rid}"):
                topic_values = [t.strip() for t in topics.splitlines() if t.strip()]
                if not topic_values:
                    st.warning("Add at least one topic.")
                else:
                    DB["Roadmap"].extend(
                        normalize_record("Roadmap", {
                            "Roadmap": rid,
                            "Heading": selected_heading,
                            "Sub": "" if selected_sub == "(none)" else selected_sub,
                            "Topic": topic, "Done": "",
                            "Order": next_order + n,
                        })
                        for n, topic in enumerate(topic_values)
                    )
                    save_db(DB)
                    _refresh()

            # ---- Delete an entire section ----
            st.divider()
            st.markdown("#### Delete section")
            delete_heading = st.selectbox(
                "Section to delete", headings, key=f"roadmap_delete_select_{rid}",
            )
            if st.button(
                f'🗑️ Delete "{delete_heading}" and its topics',
                key=f"roadmap_delete_{rid}",
            ):
                DB["Roadmap"] = [
                    r for r in DB["Roadmap"]
                    if not (
                        r.get("Roadmap") == rid
                        and str(r.get("Heading", "")).strip() == delete_heading
                    )
                ]
                save_db(DB)
                _refresh()


def study(area: str, rid: str, title: str) -> None:
    # Roadmap is intentionally closed by default. Clicking Open launches the
    # roadmap as a right-side dialog instead of consuming the whole page.
    rows = [r for r in DB["Roadmap"] if r.get("Roadmap") == rid and str(r.get("Topic", "")).strip()]
    done = sum(str(r.get("Done", "")) == "1" for r in rows)
    total = len(rows)
    pct = round(100 * done / total) if total else 0
    st.markdown(
        f'<div class="roadmap-launch"><div><div class="roadmap-launch-title">🗺️ {esc(title)} roadmap</div>'
        f'<div class="roadmap-launch-sub">{done}/{total} topics completed · {pct}% · open as a side panel</div></div></div>',
        unsafe_allow_html=True,
    )
    if st.button("↗ Open roadmap", key=f"open_roadmap_{rid}", type="primary", width="stretch"):
        roadmap_dialog(rid, title)

    st.markdown("### 📋 Tasks")
    st.caption(f"{title} tasks. Roadmap progress and task progress are tracked separately.")
    with st.expander("Show active task table", expanded=False):
        editor(
            "Tasks",
            mask=lambda d: (d["Area"] == area) & (d["Status"].astype(str) != "Done"),
            defaults={"Area": area, "Status": "To do"},
            key=f"study_{rid}",
            show=["Task", "Subject", "Topic", "Sub-topic", "Status", "Deadline", "Minutes", "Date", "Start", "End"],
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
            # Chronological within the day; untimed tasks go to the bottom.
            x = sort_tasks_df(df[(df["Date"] == day) & (df["Status"] != "Done")])

        body_parts = []
        for _, r in x.iterrows():
            if pd.notna(r["Start"]) and r["Start"]:
                time_text = str(r["Start"])[:5]
                if pd.notna(r["End"]) and r["End"]:
                    time_text += f"–{str(r['End'])[:5]}"
            else:
                time_text = "any time"
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


def cached_excel() -> bytes | None:
    """Build the backup only when the data changed (not on every rerun)."""
    version = st.session_state.get("_db_version", 0)
    if st.session_state.get("_xlsx_ver") != version:
        try:
            st.session_state["_xlsx_bytes"] = export_excel()
        except Exception as exc:  # e.g. openpyxl not installed
            st.session_state["_xlsx_bytes"] = None
            st.session_state["_xlsx_error"] = str(exc)
        st.session_state["_xlsx_ver"] = version
    return st.session_state.get("_xlsx_bytes")


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
    st.markdown(
        '<div class="brand"><span class="brand-mark">✓</span><span>Tracker</span></div>',
        unsafe_allow_html=True,
    )
    st.caption("Your colorful personal workspace")
    page = st.radio(
        "Go to",
        list(PAGES),
        key="page",
        label_visibility="collapsed",
    )
    st.divider()

    display = AUTH_USERS[USER]["display"]
    initial = display[:1].upper()
    role_text = "Shared workspace" if USER == "admin" else "Personal workspace"
    st.markdown(
        f'<div class="sidebar-user"><span class="sidebar-avatar">{esc(initial)}</span>'
        f'<b>{esc(display)}</b><div class="sidebar-role">{esc(role_text)} · session active</div></div>',
        unsafe_allow_html=True,
    )

    if st.button("↳  Log out", width="stretch"):
        st.session_state.clear()
        st.rerun()

    xlsx = cached_excel()
    if xlsx:
        st.download_button(
            "↓  Export Excel backup",
            data=xlsx,
            file_name=f"tracker_{USER}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )
    else:
        st.caption("Excel export unavailable: add 'openpyxl' to requirements.txt.")
    st.caption("☁  Supabase cloud storage")


if st.session_state.pop("_blocked_notice", False):
    st.toast(
        "Shared admin records are read-only for your account, so those changes were not saved.",
        icon="🔒",
    )

PAGES[page]()
