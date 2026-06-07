"""
automation.py — Standalone Playwright script.
Called by app.py as a subprocess. Communicates via:
  - A JSONL log file  (progress updates → Streamlit reads this)
  - The Excel file    (updated in-place after each student)
"""

import argparse
import json
import os
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


import pandas as pd
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

load_dotenv()

# ── Column names (must match what app.py normalises to) ──
NAME_COL     = "name"
EMAIL_COL    = "email"
DONE_COL     = "completed"
DEFAULT_PASS = "V2VEDTECH"


# ─────────────────────────────────────────────
# Logging helper — appends one JSON line to the
# log file so Streamlit can tail it in real time
# ─────────────────────────────────────────────
def log(log_path: str, type_: str, msg: str):
    entry = json.dumps({"type": type_, "msg": msg})
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(entry + "\n")
    print(msg, flush=True)


# ─────────────────────────────────────────────
# Playwright: update one student
# ─────────────────────────────────────────────
def update_student(page, name: str, email: str) -> tuple[bool, str]:
    try:
        page.get_by_role("link", name="All").click()
        page.wait_for_load_state("networkidle")

        sb = page.get_by_role("searchbox", name="Search:")
        sb.click()
        sb.fill(name)
        page.wait_for_timeout(1000)

        edit_btn = page.get_by_title("Edit")
        if edit_btn.count() == 0:
            return False, "Student not found in portal"

        edit_btn.first.click()
        page.wait_for_load_state("networkidle")

        ef = page.get_by_role("textbox", name="student@example.com")
        ef.click()
        ef.fill(email)
        print(f"Updating student: {name} → {email}")

        pf = page.locator('input[name="password"]')
        pf.click()
        pf.fill("V2VEDTECH")
        print(f"Set default password: {DEFAULT_PASS}")
        
        page.get_by_role("button", name="Update").click()
        page.wait_for_load_state("networkidle")

        return True, "Updated"
    except Exception as exc:
        return False, str(exc)


# ─────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--excel",   required=True, help="Path to the Excel file")
    parser.add_argument("--indices", required=True, help="Comma-separated row indices to process")
    parser.add_argument("--log",     required=True, help="Path to the JSONL progress log file")
    args = parser.parse_args()

    excel_path = args.excel
    indices    = [int(i) for i in args.indices.split(",") if i.strip()]
    log_path   = args.log

    # Clear/create the log file
    open(log_path, "w").close()

    # Load Excel
    df = pd.read_excel(excel_path, engine="openpyxl")
    df.columns = [str(c).strip().lower() for c in df.columns]
    if DONE_COL not in df.columns:
        df[DONE_COL] = False
    df[DONE_COL] = df[DONE_COL].fillna(False).astype(bool)

    try:
        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=False,slow_mo=100)
            ctx     = browser.new_context()
            page    = ctx.new_page()

            # ── Login ──
            page.goto("https://v2vedtech.com/login")
            page.locator("#email").fill(os.getenv("WEBSITE_EMAIL", ""))
            page.locator("#password").fill(os.getenv("WEBSITE_PASSWORD", ""))
            page.get_by_role("button", name="Login").click()
            page.wait_for_load_state("networkidle")
            log(log_path, "success", "🔑 Logged in successfully")

            # ── Navigate ──
            page.get_by_role("button", name=" Attendance ").click()
            page.get_by_role("link", name=" Internship").click()
            page.get_by_role("link", name="All").click()
            page.wait_for_timeout(800)  # wait a bit for the table to load
            # page.wait_for_load_state("networkidle")
            log(log_path, "info", "📍 Navigated to student list")

            # ── Process each student ──
            total = len(indices)
            for i, idx in enumerate(indices):
                if idx not in df.index:
                    log(log_path, "error", f"❌ Row index {idx} not found in Excel")
                    continue

                name  = str(df.at[idx, NAME_COL]).strip()
                email = str(df.at[idx, EMAIL_COL]).strip()

                log(log_path, "info", f"[{i+1}/{total}] Processing → {name}")

                ok, msg = update_student(page, name, email)

                if ok:
                    df.at[idx, DONE_COL] = True
                    log(log_path, "success", f"✅ {name}  ·  {email}")
                else:
                    log(log_path, "error", f"❌ {name}  —  {msg}")

                # Save Excel after EVERY student so progress survives a crash
                df.to_excel(excel_path, index=False, engine="openpyxl")

            ctx.close()
            browser.close()
            log(log_path, "done", f"🎉 Finished — {total} student(s) processed")

    except Exception as exc:
        log(log_path, "error", f"💥 Fatal error: {exc}")
        sys.exit(1)


if __name__ == "__main__":
    main()