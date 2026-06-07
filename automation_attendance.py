import re
import os
from urllib.parse import quote_plus

import pandas as pd
from dotenv import load_dotenv
from playwright.sync_api import Playwright, sync_playwright


load_dotenv()

EXCEL_FILE = "Fullstack offline student list - 07th June 2026 - Copy.xlsx"

NAME_COL = "name"      # change this to your actual column name
DATE_COL = "date"              # change this to your actual column name
ATTENDANCE_COL = "P or A"  # change this to your actual column name
STATUS_COL = "Status"


df = pd.read_excel(EXCEL_FILE, engine="openpyxl")

# If status column does not exist, create it
if STATUS_COL not in df.columns:
    df[STATUS_COL] = ""


def format_date(value = "07-06-2026"):
    """
    Convert Excel date to dd-mm-yyyy format.
    """
    if pd.isna(value):
        return ""

    date_value = pd.to_datetime(value, errors="coerce")

    if pd.isna(date_value):
        return str(value)

    return date_value.strftime("%d-%m-%Y")


def login(page):
    page.goto("https://v2vedtech.com/login")

    page.locator("#email").fill(os.getenv("WEBSITE_EMAIL"))
    page.locator("#password").fill(os.getenv("WEBSITE_PASSWORD"))

    page.get_by_role("button", name="Login").click()

    # Optional wait after login
    page.wait_for_load_state("networkidle")

def update_attendance(page, student_name: str, date: str, attendance: str) -> None:
    search_name = quote_plus(student_name)
    print(f"Processing: {student_name} | Excel Status: {attendance}")  

    page.goto(
        f"https://v2vedtech.com/admin/internship/attendance"
        f"?date={date}&stream_id=&search={search_name}"
    )

    page.wait_for_load_state("networkidle")

    # 1. If Excel says "P", skip the web elements and return immediately.
    # The main loop will automatically mark this row as "Done" in Excel.
    if attendance.lower() == "p":
        print(f"-> Skipping {student_name} (Already 'P')")
        return 

    # 2. If Excel says "A" (or anything else), change them to "P" on the portal.
    else:
        print(f"-> Changing {student_name} from 'A' to 'P' on portal...")
        
        # Using a slightly safer locator strategy to ensure we hit the right element
        attendance_button = page.get_by_text("P", exact=True).first.click()
              
        page.get_by_role("button", name=" Save Attendance").click()
        page.wait_for_load_state("networkidle")
        print(f"-> Attendance saved successfully for {student_name}")
    # else:
    #         raise Exception(f"The 'P' selection target was not visible on the page.")

def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()

    try:
        login(page)

        for index, row in df.iterrows():
            current_status = str(row.get(STATUS_COL, "")).strip().lower()

            # Skip already completed rows
            if current_status == "done":
                continue

            student_name = str(row[NAME_COL]).strip()
            date = format_date(row[DATE_COL])
            attendance = str(row[ATTENDANCE_COL]).strip()

            try:
                update_attendance(
                    page=page,
                    student_name=student_name,
                    date="07-06-2026",
                    attendance=attendance
                )

                # Mark as Done after successful update
                df.at[index, STATUS_COL] = "Done"

            except Exception as e:
                # Optional: save error in status
                df.at[index, STATUS_COL] = f"Failed: {e}"

            # Save after every row so progress is not lost
            df.to_excel(EXCEL_FILE, index=False, engine="openpyxl")

    finally:
        context.close()
        browser.close()


with sync_playwright() as playwright:
    run(playwright)