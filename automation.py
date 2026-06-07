import re
from playwright.sync_api import Playwright, sync_playwright, expect
import os
from dotenv import load_dotenv
import pandas as pd

load_dotenv()

EXCEL_FILE = "Fullstack_python_offline_students.xlsx"

df = pd.read_excel(EXCEL_FILE, engine='openpyxl')

# Add a 'Completed' column if it doesn't exist
if 'Completed' not in df.columns:
    df['Completed'] = False


def update_student_email(page, student_name: str, student_email: str) -> bool:
    try:
        # Go back to the listing page (in case we navigated away)
        page.get_by_role("link", name="All").click()
        page.wait_for_load_state("networkidle")

        # Search for the student
        searchbox = page.get_by_role("searchbox", name="Search:")
        searchbox.click()
        searchbox.fill(student_name)
        page.wait_for_timeout(1000)  # Wait for search results to filter

        # Check if student was found
        edit_btn = page.get_by_title("Edit")
        if edit_btn.count() == 0:
            print(f"  ⚠️  No record found for: {student_name}")
            return False

        edit_btn.first.click()
        page.wait_for_load_state("networkidle")

        # Update email
        email_field = page.get_by_role("textbox", name="student@example.com")
        email_field.click()
        email_field.fill(student_email)

        # Set default password
        pwd_field = page.get_by_role("textbox", name="Leave blank to skip")
        pwd_field.click()
        pwd_field.fill("V2VEDTECH")

        page.get_by_role("button", name="Update").click()
        page.wait_for_load_state("networkidle")

        print(f"  ✅ Updated: {student_name} → {student_email}")
        return True

    except Exception as e:
        print(f"  ❌ Error updating {student_name}: {e}")
        return False


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()

    # --- Login ---
    page.goto("https://v2vedtech.com/login")
    page.locator("#email").fill(os.getenv("WEBSITE_EMAIL"))
    page.locator("#password").fill(os.getenv("WEBSITE_PASSWORD"))
    page.get_by_role("button", name="Login").click()
    page.wait_for_load_state("networkidle")

    # --- Navigate to Attendance > Internship > All ---
    page.get_by_role("button", name=" Attendance ").click()
    page.get_by_role("link", name=" Internship").click()
    page.get_by_role("link", name="All").click()
    page.wait_for_load_state("networkidle")

    # --- Loop through students ---
    # Adjust column names below to match your actual Excel headers
    NAME_COL = "Student Name"   # ← change if different
    EMAIL_COL = "Email"         # ← change if different

    total = len(df)
    for idx, row in df.iterrows():

        # Skip already completed rows
        if row.get("Completed") == True:
            print(f"[{idx+1}/{total}] Skipping (already done): {row[NAME_COL]}")
            continue

        student_name = str(row[NAME_COL]).strip()
        student_email = str(row[EMAIL_COL]).strip()

        print(f"[{idx+1}/{total}] Processing: {student_name}")

        success = update_student_email(page, student_name, student_email)

        if success:
            df.at[idx, 'Completed'] = True
            # Save after every successful update so progress is never lost
            df.to_excel(EXCEL_FILE, index=False, engine='openpyxl')

    print("\n🎉 All done!")
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)