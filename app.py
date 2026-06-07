# """
# app.py — Streamlit admin UI for V2VedTech email updates.
# Spawns automation.py as a subprocess to avoid asyncio conflicts with Playwright.
# """

# import json
# import os
# import subprocess
# import sys
# import tempfile
# import time
# from io import BytesIO
# from pathlib import Path

# import pandas as pd
# import streamlit as st

# # ─────────────────────────────────────────────
# # Constants
# # ─────────────────────────────────────────────
# NAME_COL = "name"
# EMAIL_COL = "email"
# DONE_COL  = "completed"

# # Temp dir for Excel + log files shared with subprocess
# WORK_DIR  = Path(tempfile.gettempdir()) / "v2vedtech"
# WORK_DIR.mkdir(exist_ok=True)

# EXCEL_PATH = WORK_DIR / "students.xlsx"
# LOG_PATH   = WORK_DIR / "progress.jsonl"

# # ─────────────────────────────────────────────
# # Page config
# # ─────────────────────────────────────────────
# st.set_page_config(
#     page_title="V2VedTech Admin",
#     page_icon="🎓",
#     layout="wide",
#     initial_sidebar_state="expanded",
# )

# # ─────────────────────────────────────────────
# # CSS
# # ─────────────────────────────────────────────
# st.markdown("""
# <style>
# @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=DM+Mono:wght@400;500&display=swap');

# html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }

# [data-testid="stSidebar"] {
#     background: linear-gradient(180deg, #0f172a 0%, #1e293b 100%);
#     border-right: 1px solid #334155;
# }
# [data-testid="stSidebar"] * { color: #e2e8f0 !important; }
# [data-testid="stSidebar"] .stRadio > label { display: none; }

# [data-testid="stMetric"] {
#     background: #f8fafc;
#     border: 1px solid #e2e8f0;
#     border-radius: 12px;
#     padding: 1rem 1.25rem;
# }
# [data-testid="stMetricValue"] { font-family: 'DM Mono', monospace; }
# [data-testid="stMetric"] * { color: #000000 !important; }

# .page-title-white,
# .section-title-white,
# .notice-text-white {
#     color: #ffffff !important;
# }

# .stButton > button {
#     font-family: 'DM Sans', sans-serif;
#     font-weight: 600;
#     border-radius: 8px;
#     transition: all 0.2s ease;
# }
# .stButton > button[kind="primary"] {
#     background: linear-gradient(135deg, #6366f1, #8b5cf6);
#     border: none;
#     color: white;
# }
# .stButton > button[kind="primary"]:hover {
#     background: linear-gradient(135deg, #4f46e5, #7c3aed);
#     transform: translateY(-1px);
#     box-shadow: 0 4px 12px rgba(99,102,241,0.4);
# }

# [data-testid="stFileUploader"] {
#     border: 2px dashed #cbd5e1;
#     border-radius: 12px;
#     padding: 1rem;
#     background: #f8fafc;
# }

# .log-box {
#     background: #0f172a;
#     border-radius: 10px;
#     padding: 1rem 1.25rem;
#     max-height: 340px;
#     overflow-y: auto;
#     font-family: 'DM Mono', monospace;
#     font-size: 0.82rem;
#     border: 1px solid #1e293b;
# }
# .log-entry-success { color: #4ade80; }
# .log-entry-error   { color: #f87171; }
# .log-entry-info    { color: #60a5fa; }
# .log-entry-done    { color: #facc15; font-weight: 600; }

# h2 { color: #0f172a !important; font-weight: 700 !important; }
# h3 { color: #1e293b !important; font-weight: 600 !important; }

# .stDownloadButton > button {
#     background: #10b981 !important;
#     color: white !important;
#     border: none !important;
#     font-weight: 600 !important;
#     border-radius: 8px !important;
# }
# .coming-soon {
#     display: flex; flex-direction: column; align-items: center;
#     justify-content: center; padding: 5rem 2rem; text-align: center; color: #64748b;
# }
# .coming-soon .icon { font-size: 5rem; margin-bottom: 1.5rem; }
# </style>
# """, unsafe_allow_html=True)


# # ─────────────────────────────────────────────
# # Session state init
# # ─────────────────────────────────────────────
# for key, default in {
#     "df"          : None,
#     "filename"    : None,
#     "proc"        : None,   # subprocess.Popen object
#     "is_running"  : False,
# }.items():
#     if key not in st.session_state:
#         st.session_state[key] = default


# # ─────────────────────────────────────────────
# # Helpers
# # ─────────────────────────────────────────────
# def save_excel_to_disk(df: pd.DataFrame):
#     """Write current df to the shared temp Excel so the subprocess can read it."""
#     df.to_excel(EXCEL_PATH, index=False, engine="openpyxl")


# def reload_excel_from_disk() -> pd.DataFrame:
#     """Read the Excel back after subprocess updates it."""
#     df = pd.read_excel(EXCEL_PATH, engine="openpyxl")
#     df.columns = [str(c).strip().lower() for c in df.columns]
#     if DONE_COL not in df.columns:
#         df[DONE_COL] = False
#     df[DONE_COL] = df[DONE_COL].fillna(False).astype(bool)
#     return df


# def read_log() -> list[dict]:
#     """Read all JSONL lines from the progress log."""
#     if not LOG_PATH.exists():
#         return []
#     lines = []
#     with open(LOG_PATH, encoding="utf-8") as f:
#         for line in f:
#             line = line.strip()
#             if line:
#                 try:
#                     lines.append(json.loads(line))
#                 except json.JSONDecodeError:
#                     pass
#     return lines


# def spawn_automation(indices: list[int]):
#     """Launch automations.py as a completely separate process."""
#     LOG_PATH.unlink(missing_ok=True)   # clear old log

#     env = os.environ.copy()
#     env["PYTHONIOENCODING"] = "utf-8"

#     proc = subprocess.Popen(
#         [
#             sys.executable,            # same Python interpreter
#             str(Path(__file__).parent / "automations.py"),
#             "--excel",   str(EXCEL_PATH),
#             "--indices", ",".join(str(i) for i in indices),
#             "--log",     str(LOG_PATH),
#         ],
#         stdout=subprocess.PIPE,
#         stderr=subprocess.STDOUT,
#     )
#     st.session_state.proc       = proc
#     st.session_state.is_running = True


# def check_proc_alive() -> bool:
#     """Return True if the subprocess is still running."""
#     proc = st.session_state.get("proc")
#     if proc is None:
#         return False
#     return proc.poll() is None   # None means still running


# def render_log(entries: list[dict]):
#     """Render progress log as a dark terminal box."""
#     if not entries:
#         return
#     lines = []
#     for e in entries:
#         cls = {
#             "success": "log-entry-success",
#             "error"  : "log-entry-error",
#             "done"   : "log-entry-done",
#         }.get(e.get("type", ""), "log-entry-info")
#         msg = e.get("msg", "")
#         lines.append(f'<div class="{cls}">{msg}</div>')

#     st.markdown(
#         f'<div class="log-box">{"".join(reversed(lines))}</div>',
#         unsafe_allow_html=True,
#     )


# # ─────────────────────────────────────────────
# # Poll: if subprocess just finished, reload df
# # ─────────────────────────────────────────────
# if st.session_state.is_running and not check_proc_alive():
#     # Subprocess finished — reload updated Excel into session
#     st.session_state.is_running = False
#     st.session_state.proc       = None
#     if EXCEL_PATH.exists():
#         st.session_state.df = reload_excel_from_disk()


# # ─────────────────────────────────────────────
# # Sidebar
# # ─────────────────────────────────────────────
# with st.sidebar:
#     st.markdown("## 🎓 V2VedTech")
#     st.markdown("<hr style='border-color:#334155;margin:0.5rem 0 1rem'>", unsafe_allow_html=True)
#     nav = st.radio(
#         "nav",
#         ["📋  Attendance", "📧  Email Attendance"],
#         label_visibility="collapsed",
#     )
#     st.markdown("<br>", unsafe_allow_html=True)
#     st.markdown(
#         "<div style='font-size:0.75rem;color:#475569;padding:0 0.25rem'>"
#         "V2VedTech Admin Portal · v1.0</div>",
#         unsafe_allow_html=True,
#     )


# # ═════════════════════════════════════════════
# # PAGE — Attendance (coming soon)
# # ═════════════════════════════════════════════
# if nav == "📋  Attendance":
#     st.markdown("""
#     <div class="coming-soon">
#         <div class="icon">🚧</div>
#         <h2>Attendance — Coming Soon</h2>
#         <p>This feature is currently under development. Check back soon!</p>
#     </div>
#     """, unsafe_allow_html=True)


# # ═════════════════════════════════════════════
# # PAGE — Email Attendance
# # ═════════════════════════════════════════════
# else:
#     st.markdown("## 📧 Email Attendance")
#     st.markdown(
#         "<p style='color:#64748b;margin-top:-0.5rem;margin-bottom:1.5rem'>"
#         "Upload the student roster, fill missing emails, and sync to the portal automatically.</p>",
#         unsafe_allow_html=True,
#     )

#     # ── File uploader ────────────────────────
#     uploaded = st.file_uploader("Upload student Excel file (.xlsx)", type=["xlsx"])

#     if uploaded:
#         # Reload only when the file changes
#         if (
#             st.session_state.df is None
#             or st.session_state.filename != uploaded.name
#         ):
#             raw = pd.read_excel(uploaded, engine="openpyxl")
#             raw.columns = [str(c).strip().lower() for c in raw.columns]
#             if DONE_COL not in raw.columns:
#                 raw[DONE_COL] = False
#             raw[DONE_COL] = raw[DONE_COL].fillna(False).astype(bool)
#             st.session_state.df       = raw
#             st.session_state.filename = uploaded.name

#             # Write to disk so subprocess can access it
#             save_excel_to_disk(raw)
#             st.success(f"✅ Loaded **{uploaded.name}** — {len(raw)} students found")

#         df = st.session_state.df

#         # ── Stats ────────────────────────────
#         total         = len(df)
#         completed     = int(df[DONE_COL].sum())
#         has_email     = df[EMAIL_COL].notna() & (df[EMAIL_COL].astype(str).str.strip() != "")
#         missing_count = int((~has_email).sum())
#         pending       = total - completed

#         st.markdown("<br>", unsafe_allow_html=True)
#         c1, c2, c3, c4 = st.columns(4)
#         c1.metric("👥 Total Students", total)
#         c2.metric("✅ Completed",       completed)
#         c3.metric("⏳ Pending",         pending)
#         c4.metric("❓ Missing Email",   missing_count)

#         st.markdown("---")

#         # ── Automation ───────────────────────
#         st.markdown("### 🤖 Run Automation")

#         ready_df = df[
#             (~df[DONE_COL]) &
#             df[EMAIL_COL].notna() &
#             (df[EMAIL_COL].astype(str).str.strip() != "")
#         ]

#         if len(ready_df) == 0:
#             st.info("No students with emails are pending an update right now.")
#         else:
#             st.markdown(
#                 f"<p style='color:#475569'><b>{len(ready_df)}</b> student(s) ready to be updated.</p>",
#                 unsafe_allow_html=True,
#             )

#             if not st.session_state.is_running:
#                 if st.button("▶️  Start Update", type="primary", use_container_width=True):
#                     save_excel_to_disk(st.session_state.df)   # write latest df to disk
#                     spawn_automation(ready_df.index.tolist())
#                     st.rerun()
#             else:
#                 st.warning("⏳ Automation is running — keep this tab open.")
#                 col_r, col_s = st.columns(2)
#                 with col_r:
#                     if st.button("🔄  Refresh Progress", use_container_width=True):
#                         st.rerun()
#                 with col_s:
#                     if st.button("🛑  Stop", use_container_width=True):
#                         proc = st.session_state.get("proc")
#                         if proc:
#                             proc.terminate()
#                         st.session_state.is_running = False
#                         st.session_state.proc       = None
#                         if EXCEL_PATH.exists():
#                             st.session_state.df = reload_excel_from_disk()
#                         st.rerun()

#         # ── Progress log ─────────────────────
#         log_entries = read_log()
#         if log_entries:
#             st.markdown("#### 📊 Live Progress")
#             render_log(log_entries)

#         # Auto-refresh every 2 s while subprocess is alive
#         if st.session_state.is_running:
#             time.sleep(2)
#             st.rerun()

#         st.markdown("---")

#         # ── Editable table for missing emails ─
#         st.markdown("### ✏️ Fill Missing Emails")

#         missing_mask = ~(
#             df[EMAIL_COL].notna() &
#             (df[EMAIL_COL].astype(str).str.strip() != "")
#         )
#         missing_df = df[missing_mask].copy()

#         if missing_df.empty:
#             st.success("✅ All students have email addresses — nothing to fill!")
#         else:
#             st.markdown(
#                 f"<p style='color:#475569'>"
#                 f"<b>{len(missing_df)}</b> student(s) are missing an email. "
#                 f"Edit the <b>email</b> column below and click <b>Save &amp; Update</b>.</p>",
#                 unsafe_allow_html=True,
#             )

#             # Column config — only email is editable
#             col_cfg = {}
#             for col in missing_df.columns:
#                 if col == EMAIL_COL:
#                     col_cfg[col] = st.column_config.TextColumn(
#                         "Email ✏️",
#                         help="Enter the student's email address",
#                         validate=r"^[^@\s]+@[^@\s]+\.[^@\s]+$",
#                     )
#                 elif col == DONE_COL:
#                     col_cfg[col] = st.column_config.CheckboxColumn("Done ✓", disabled=True)
#                 else:
#                     col_cfg[col] = st.column_config.TextColumn(col.title(), disabled=True)

#             edited = st.data_editor(
#                 missing_df,
#                 column_config=col_cfg,
#                 use_container_width=True,
#                 hide_index=True,
#                 num_rows="fixed",
#                 key="missing_editor",
#             )

#             if st.button("💾  Save & Update Missing", type="primary", use_container_width=True):
#                 newly_filled = []
#                 for idx in edited.index:
#                     val = str(edited.at[idx, EMAIL_COL]).strip()
#                     if val and val.lower() not in ("nan", "none", ""):
#                         st.session_state.df.at[idx, EMAIL_COL] = val
#                         newly_filled.append(idx)

#                 if newly_filled:
#                     save_excel_to_disk(st.session_state.df)
#                     spawn_automation(newly_filled)
#                     st.rerun()
#                 else:
#                     st.warning("⚠️ No emails were entered. Please fill at least one row.")

#         st.markdown("---")

#         # ── Download ─────────────────────────
#         st.markdown("### ⬇️ Download Updated File")
#         st.markdown(
#             "<p style='color:#475569;margin-top:-0.5rem'>"
#             "Downloads the latest version with updated emails and completion status.</p>",
#             unsafe_allow_html=True,
#         )

#         buf = BytesIO()
#         st.session_state.df.to_excel(buf, index=False, engine="openpyxl")
#         buf.seek(0)

#         st.download_button(
#             label="📥  Download Excel",
#             data=buf,
#             file_name=f"updated_{st.session_state.filename or 'students.xlsx'}",
#             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#             use_container_width=True,
#         )

#     else:
#         # Empty state
#         st.markdown("""
#         <div style="
#             text-align:center;padding:4rem 2rem;
#             border:2px dashed #cbd5e1;border-radius:16px;
#             background:#f8fafc;color:#94a3b8;margin-top:1rem
#         ">
#             <div style="font-size:3.5rem;margin-bottom:1rem">📂</div>
#             <div style="font-size:1.1rem;font-weight:600;color:#475569">
#                 Upload your student Excel file to get started
#             </div>
#             <div style="font-size:0.9rem;margin-top:0.5rem">
#                 Expected columns: id · name · number · location · course · mode · email · status
#             </div>
#         </div>
#         """, unsafe_allow_html=True)

"""
app.py — Streamlit admin UI for V2VedTech email updates.
Spawns automation.py as a subprocess to avoid asyncio conflicts with Playwright.
"""

import json
import os
import re
import subprocess
import sys
import tempfile
import time
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

# ─────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────
NAME_COL = "name"
EMAIL_COL = "email"
DONE_COL  = "completed"

# Temp dir for Excel + log files shared with subprocess
WORK_DIR  = Path(tempfile.gettempdir()) / "v2vedtech"
WORK_DIR.mkdir(exist_ok=True)

EXCEL_PATH = WORK_DIR / "students.xlsx"
LOG_PATH   = WORK_DIR / "progress.jsonl"

# ─────────────────────────────────────────────
# Page config
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="V2VedTech Admin",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=DM+Mono:wght@400;500&display=swap');

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }

[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0f172a 0%, #1e293b 100%);
    border-right: 1px solid #334155;
}
[data-testid="stSidebar"] * { color: #e2e8f0 !important; }
[data-testid="stSidebar"] .stRadio > label { display: none; }

[data-testid="stMetric"] {
    background: #f8fafc;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 1rem 1.25rem;
}
[data-testid="stMetricValue"] { font-family: 'DM Mono', monospace; }
[data-testid="stMetric"] * { color: #000000 !important; }

.page-title-white,
.section-title-white,
.notice-text-white {
    color: #ffffff !important;
}

.stButton > button {
    font-family: 'DM Sans', sans-serif;
    font-weight: 600;
    border-radius: 8px;
    transition: all 0.2s ease;
}
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #6366f1, #8b5cf6);
    border: none;
    color: white;
}
.stButton > button[kind="primary"]:hover {
    background: linear-gradient(135deg, #4f46e5, #7c3aed);
    transform: translateY(-1px);
    box-shadow: 0 4px 12px rgba(99,102,241,0.4);
}

[data-testid="stFileUploader"] {
    border: 2px dashed #cbd5e1;
    border-radius: 12px;
    padding: 1rem;
    background: #f8fafc;
}

.log-box {
    background: #0f172a;
    border-radius: 10px;
    padding: 1rem 1.25rem;
    max-height: 340px;
    overflow-y: auto;
    font-family: 'DM Mono', monospace;
    font-size: 0.82rem;
    border: 1px solid #1e293b;
}
.log-entry-success { color: #4ade80; }
.log-entry-error   { color: #f87171; }
.log-entry-info    { color: #60a5fa; }
.log-entry-done    { color: #facc15; font-weight: 600; }

h2 { color: #0f172a !important; font-weight: 700 !important; }
h3 { color: #1e293b !important; font-weight: 600 !important; }

.stDownloadButton > button {
    background: #10b981 !important;
    color: white !important;
    border: none !important;
    font-weight: 600 !important;
    border-radius: 8px !important;
}
.coming-soon {
    display: flex; flex-direction: column; align-items: center;
    justify-content: center; padding: 5rem 2rem; text-align: center; color: #64748b;
}
.coming-soon .icon { font-size: 5rem; margin-bottom: 1.5rem; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# Session state init
# ─────────────────────────────────────────────
for key, default in {
    "df"          : None,
    "filename"    : None,
    "proc"        : None,   # subprocess.Popen object
    "is_running"  : False,
}.items():
    if key not in st.session_state:
        st.session_state[key] = default


# ─────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────
def save_excel_to_disk(df: pd.DataFrame):
    """Write current df to the shared temp Excel so the subprocess can read it."""
    df.to_excel(EXCEL_PATH, index=False, engine="openpyxl")


def reload_excel_from_disk() -> pd.DataFrame:
    """Read the Excel back after subprocess updates it."""
    df = pd.read_excel(EXCEL_PATH, engine="openpyxl")
    df.columns = [str(c).strip().lower() for c in df.columns]
    if DONE_COL not in df.columns:
        df[DONE_COL] = False
    df[DONE_COL] = df[DONE_COL].fillna(False).astype(bool)
    return df


def read_log() -> list[dict]:
    """Read all JSONL lines from the progress log."""
    if not LOG_PATH.exists():
        return []
    lines = []
    with open(LOG_PATH, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                try:
                    lines.append(json.loads(line))
                except json.JSONDecodeError:
                    pass
    return lines


def spawn_automation(indices: list[int]):
    """Launch automations.py as a completely separate process."""
    LOG_PATH.unlink(missing_ok=True)   # clear old log

    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"

    proc = subprocess.Popen(
        [
            sys.executable,            # same Python interpreter
            str(Path(__file__).parent / "automations.py"),
            "--excel",   str(EXCEL_PATH),
            "--indices", ",".join(str(i) for i in indices),
            "--log",     str(LOG_PATH),
        ],
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
    )
    st.session_state.proc       = proc
    st.session_state.is_running = True


def check_proc_alive() -> bool:
    """Return True if the subprocess is still running."""
    proc = st.session_state.get("proc")
    if proc is None:
        return False
    return proc.poll() is None   # None means still running


def render_log(entries: list[dict]):
    """Render progress log as a dark terminal box."""
    if not entries:
        return
    lines = []
    for e in entries:
        cls = {
            "success": "log-entry-success",
            "error"  : "log-entry-error",
            "done"   : "log-entry-done",
        }.get(e.get("type", ""), "log-entry-info")
        msg = e.get("msg", "")
        lines.append(f'<div class="{cls}">{msg}</div>')

    st.markdown(
        f'<div class="log-box">{"".join(reversed(lines))}</div>',
        unsafe_allow_html=True,
    )


# ─────────────────────────────────────────────
# Poll: if subprocess just finished, reload df
# ─────────────────────────────────────────────
if st.session_state.is_running and not check_proc_alive():
    # Subprocess finished — reload updated Excel into session
    st.session_state.is_running = False
    st.session_state.proc       = None
    if EXCEL_PATH.exists():
        st.session_state.df = reload_excel_from_disk()


# ─────────────────────────────────────────────
# Sidebar
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🎓 V2VedTech")
    st.markdown("<hr style='border-color:#334155;margin:0.5rem 0 1rem'>", unsafe_allow_html=True)
    nav = st.radio(
        "nav",
        ["📋  Attendance", "📧  Email Updation"],
        label_visibility="collapsed",
    )
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        "<div style='font-size:0.75rem;color:#475569;padding:0 0.25rem'>"
        "V2VedTech Admin Portal · v1.0</div>",
        unsafe_allow_html=True,
    )


# ═════════════════════════════════════════════
# PAGE — Attendance (coming soon)
# ═════════════════════════════════════════════
if nav == "📋  Attendance":
    st.markdown("""
    <div class="coming-soon">
        <div class="icon">🚧</div>
        <h2>Attendance — Coming Soon</h2>
        <p>This feature is currently under development. Check back soon!</p>
    </div>
    """, unsafe_allow_html=True)


# ═════════════════════════════════════════════
# PAGE — Email Updation
# ═════════════════════════════════════════════
else:
    st.markdown("<h2 class='page-title-white'>📧 Email Updation</h2>", unsafe_allow_html=True)
    st.markdown(
        "<p style='color:#64748b;margin-top:-0.5rem;margin-bottom:1.5rem'>"
        "Upload the student roster, fill missing emails, and sync to the portal automatically.</p>",
        unsafe_allow_html=True,
    )

    # ── File uploader ────────────────────────
    uploaded = st.file_uploader("Upload student Excel file (.xlsx)", type=["xlsx"])

    if uploaded:
        # Reload only when the file changes
        if (
            st.session_state.df is None
            or st.session_state.filename != uploaded.name
        ):
            raw = pd.read_excel(uploaded, engine="openpyxl")
            raw.columns = [str(c).strip().lower() for c in raw.columns]
            if DONE_COL not in raw.columns:
                raw[DONE_COL] = False
            raw[DONE_COL] = raw[DONE_COL].fillna(False).astype(bool)
            st.session_state.df       = raw
            st.session_state.filename = uploaded.name

            # Write to disk so subprocess can access it
            save_excel_to_disk(raw)
            st.success(f"✅ Loaded **{uploaded.name}** — {len(raw)} students found")

        df = st.session_state.df

        # ── Stats ────────────────────────────
        total         = len(df)
        completed     = int(df[DONE_COL].sum())
        has_email     = df[EMAIL_COL].notna() & (df[EMAIL_COL].astype(str).str.strip() != "")
        missing_count = int((~has_email).sum())
        pending       = total - completed

        st.markdown("<br>", unsafe_allow_html=True)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("👥 Total Students", total)
        c2.metric("✅ Completed",       completed)
        c3.metric("⏳ Pending",         pending)
        c4.metric("❓ Missing Email",   missing_count)

        st.markdown("---")

        # ── Automation ───────────────────────
        st.markdown("<h3 class='section-title-white'>🤖 Run Automation</h3>", unsafe_allow_html=True)

        ready_df = df[
            (~df[DONE_COL]) &
            df[EMAIL_COL].notna() &
            (df[EMAIL_COL].astype(str).str.strip() != "")
        ]

        if len(ready_df) == 0:
            st.info("No students with emails are pending an update right now.")
        else:
            st.markdown(
                f"<p style='color:#475569'><b>{len(ready_df)}</b> student(s) ready to be updated.</p>",
                unsafe_allow_html=True,
            )

            if not st.session_state.is_running:
                if st.button("▶️  Start Update", type="primary", use_container_width=True):
                    save_excel_to_disk(st.session_state.df)   # write latest df to disk
                    spawn_automation(ready_df.index.tolist())
                    st.rerun()
            else:
                st.warning("⏳ Automation is running — keep this tab open.")
                col_r, col_s = st.columns(2)
                with col_r:
                    if st.button("🔄  Refresh Progress", use_container_width=True):
                        st.rerun()
                with col_s:
                    if st.button("🛑  Stop", use_container_width=True):
                        proc = st.session_state.get("proc")
                        if proc:
                            proc.terminate()
                        st.session_state.is_running = False
                        st.session_state.proc       = None
                        if EXCEL_PATH.exists():
                            st.session_state.df = reload_excel_from_disk()
                        st.rerun()

        # ── Progress log ─────────────────────
        log_entries = read_log()
        if log_entries:
            st.markdown("#### 📊 Live Progress")
            render_log(log_entries)

        # Auto-refresh every 2 s while subprocess is alive
        if st.session_state.is_running:
            time.sleep(2)
            st.rerun()

        st.markdown("---")

        # ── Editable table for missing emails ─
        st.markdown("<h3 class='section-title-white'>✏️ Fill Missing Emails</h3>", unsafe_allow_html=True)

        missing_mask = ~(
            df[EMAIL_COL].notna() &
            (df[EMAIL_COL].astype(str).str.strip() != "")
        )
        missing_df = df.loc[missing_mask].copy()

        if missing_df.empty:
            st.success("✅ All students have email addresses — nothing to fill!")
        else:
            st.markdown(
                f"<p class='notice-text-white'>"
                f"<b>{len(missing_df)}</b> student(s) are missing an email. "
                f"Edit the <b>email</b> column below and click <b>Save &amp; Update</b>.</p>",
                unsafe_allow_html=True,
            )

            # Keep the original dataframe row index inside the editor.
            # This makes saving reliable even after filtering missing-email rows.
            ROW_ID_COL = "_row_id"
            editor_df = missing_df.copy()
            editor_df.insert(0, ROW_ID_COL, editor_df.index)
            editor_df = editor_df.reset_index(drop=True)

            # Column config — only email is editable
            col_cfg = {
                ROW_ID_COL: st.column_config.NumberColumn(
                    "Row ID",
                    help="Original row number from the uploaded dataframe",
                    disabled=True,
                )
            }
            for col in editor_df.columns:
                if col == ROW_ID_COL:
                    continue
                if col == EMAIL_COL:
                    col_cfg[col] = st.column_config.TextColumn(
                        "Email ✏️",
                        help="Enter the student's email address, or leave it blank",
                        # Allow either a valid email OR a completely blank cell.
                        validate=r"^$|^[^@\s]+@[^@\s]+\.[^@\s]+$",
                    )
                elif col == DONE_COL:
                    col_cfg[col] = st.column_config.CheckboxColumn("Done ✓")
                else:
                    col_cfg[col] = st.column_config.TextColumn(col.title())

            disabled_cols = [col for col in editor_df.columns if col != EMAIL_COL]

            edited = st.data_editor(
                editor_df,
                column_config=col_cfg,
                disabled=disabled_cols,
                use_container_width=True,
                hide_index=True,
                num_rows="fixed",
                key="missing_editor",
            )

            if st.button("💾  Save & Update Missing", type="primary", use_container_width=True):
                rows_to_automate = []
                rows_updated = []
                rows_cleared = []
                invalid_rows = []
                email_pattern = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

                for _, row in edited.iterrows():
                    original_idx = int(row[ROW_ID_COL])

                    raw_val = row.get(EMAIL_COL, "")
                    new_val = "" if pd.isna(raw_val) else str(raw_val).strip()
                    if new_val.lower() in ("nan", "none"):
                        new_val = ""

                    old_raw = st.session_state.df.at[original_idx, EMAIL_COL]
                    old_val = "" if pd.isna(old_raw) else str(old_raw).strip()
                    if old_val.lower() in ("nan", "none"):
                        old_val = ""

                    # Blank is allowed. Save it back as blank instead of ignoring it.
                    if new_val == "":
                        st.session_state.df.at[original_idx, EMAIL_COL] = ""
                        if old_val != "":
                            rows_cleared.append(original_idx)
                            rows_updated.append(original_idx)
                        continue

                    # Non-empty values must be valid email addresses.
                    if not email_pattern.match(new_val):
                        invalid_rows.append(original_idx)
                        continue

                    st.session_state.df.at[original_idx, EMAIL_COL] = new_val
                    if new_val != old_val:
                        rows_updated.append(original_idx)

                    # Only rows with a newly entered email should go to automation.
                    if old_val == "" and not bool(st.session_state.df.at[original_idx, DONE_COL]):
                        rows_to_automate.append(original_idx)

                rows_to_automate = list(dict.fromkeys(rows_to_automate))
                rows_updated = list(dict.fromkeys(rows_updated))
                rows_cleared = list(dict.fromkeys(rows_cleared))

                if invalid_rows:
                    st.warning("⚠️ Some emails are invalid. Please enter valid emails or leave cells blank.")
                else:
                    save_excel_to_disk(st.session_state.df)

                    if rows_to_automate:
                        spawn_automation(rows_to_automate)
                        st.rerun()
                    elif rows_cleared:
                        st.success(f"✅ Cleared email for {len(rows_cleared)} student(s).")
                        st.rerun()
                    elif rows_updated:
                        st.success(f"✅ Saved email changes for {len(rows_updated)} student(s).")
                        st.rerun()
                    else:
                        st.success("✅ Saved. Blank email cells are allowed.")

        st.markdown("---")

        # ── Download ─────────────────────────
        st.markdown("<h3 class='section-title-white'>⬇️ Download Updated File</h3>", unsafe_allow_html=True)
        st.markdown(
            "<p style='color:#475569;margin-top:-0.5rem'>"
            "Downloads the latest version with updated emails and completion status.</p>",
            unsafe_allow_html=True,
        )

        buf = BytesIO()
        st.session_state.df.to_excel(buf, index=False, engine="openpyxl")
        buf.seek(0)

        st.download_button(
            label="📥  Download Excel",
            data=buf,
            file_name=f"updated_{st.session_state.filename or 'students.xlsx'}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    else:
        # Empty state
        st.markdown("""
        <div style="
            text-align:center;padding:4rem 2rem;
            border:2px dashed #cbd5e1;border-radius:16px;
            background:#f8fafc;color:#94a3b8;margin-top:1rem
        ">
            <div style="font-size:3.5rem;margin-bottom:1rem">📂</div>
            <div style="font-size:1.1rem;font-weight:600;color:#475569">
                Upload your student Excel file to get started
            </div>
            <div style="font-size:0.9rem;margin-top:0.5rem">
                Expected columns: id · name · number · location · course · mode · email · status
            </div>
        </div>
        """, unsafe_allow_html=True)

