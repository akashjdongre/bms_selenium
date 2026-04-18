import os
import shutil
import time
from datetime import datetime, timedelta

import pandas as pd
import mysql.connector

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service

# ──────────────────────────────────────────────
# CONFIG
# ──────────────────────────────────────────────

BASE_DIR         = "/app/kirtan-baldev-meghna"
DOWNLOAD_FOLDER  = os.path.join(BASE_DIR, "downloads")
PROCESSED_FOLDER = os.path.join(BASE_DIR, "processed")  # date subfolder added at runtime
LOG_FILE         = os.path.join(BASE_DIR, "seat_scheduler_log.txt")

BMS_URL      = "https://bo.bookmyshow.com/"
BMS_USER     = "kirtan.1"
BMS_PASSWORD = "Bs@12345"

COMPANY_VALUES = ["BALG"]

EVENT_TYPE = "Kirtan Collective by Baldev & Meghna"

DB_CONFIG = {
    "host":     "host.docker.internal",
    "user":     "root",
    "password": "root",
    "database": "docker_test",
}

INSERT_QUERY = """
    INSERT INTO bms_individual_ticket_mstr (
        booking_id, booking_date, event_name, location,
        event_date, ticket_quantity, seat_info, platform,
        insert_date, report_date, update_date, event_type
    ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
    ON DUPLICATE KEY UPDATE
        ticket_quantity = VALUES(ticket_quantity),
        update_date     = VALUES(update_date),
        event_type      = VALUES(event_type);
"""


# ──────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────

def log(msg: str) -> None:
    """Append a timestamped message to the log file and print it."""
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def excel_serial_to_datetime(value) -> str | None:
    """Convert an Excel serial date number to a formatted datetime string."""
    if pd.isna(value):
        return None
    try:
        return (datetime(1899, 12, 30) + timedelta(days=float(value))).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return None


def make_processed_folder() -> str:
    """Return (and create if needed) today's dated subfolder inside processed/."""
    folder = os.path.join(PROCESSED_FOLDER, datetime.now().strftime("%Y-%m-%d"))
    os.makedirs(folder, exist_ok=True)
    return folder


# ──────────────────────────────────────────────
# SELENIUM – download reports
# ──────────────────────────────────────────────

def build_driver() -> webdriver.Chrome:
    opts = Options()
    opts.add_argument("--headless")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1920,1080")
    opts.binary_location = os.environ.get("CHROME_BIN", "/usr/bin/chromium")
    service = Service(executable_path=os.environ.get("CHROMEDRIVER_PATH", "/usr/bin/chromedriver"))
    opts.add_experimental_option("prefs", {
        "download.default_directory":               DOWNLOAD_FOLDER,
        "download.prompt_for_download":             False,
        "profile.default_content_settings.popups":  0,
    })
    return webdriver.Chrome(service=service, options=opts)


def login(driver: webdriver.Chrome, wait: WebDriverWait) -> None:
    driver.get(BMS_URL)
    # driver.maximize_window()  # Not needed in headless
    time.sleep(3)

    driver.find_element(By.ID, "txtUserId").send_keys(BMS_USER)
    driver.find_element(By.ID, "txtPassword").send_keys(BMS_PASSWORD)
    driver.find_element(By.ID, "cmdLogin").click()

    time.sleep(10)
    log("Login successful")


def debug_snapshot(driver: webdriver.Chrome, label: str) -> None:
    """Save a screenshot and page source for debugging."""
    ts = datetime.now().strftime("%H%M%S")
    shot_path = os.path.join(BASE_DIR, f"debug_{label}_{ts}.png")
    html_path  = os.path.join(BASE_DIR, f"debug_{label}_{ts}.html")
    driver.save_screenshot(shot_path)
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(driver.page_source)
    log(f"Debug snapshot saved: {shot_path}")


def navigate_to_report(driver: webdriver.Chrome, wait: WebDriverWait) -> None:
    """Click through Reports -> Sales Reports -> Event Show Wise New."""
    driver.switch_to.default_content()
    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
    debug_snapshot(driver, "after_login")
    for xpath in [
        "//a[normalize-space()='Reports']",
        "//a[normalize-space()='Sales Reports']",
        "//a[normalize-space()='Event Show Wise New']",
    ]:
        log(f"Clicking on element with XPath: {xpath}")
        try:
            wait.until(EC.element_to_be_clickable((By.XPATH, xpath))).click()
        except Exception:
            # Nav link may be inside a frame — try switching into the first iframe
            driver.switch_to.default_content()
            debug_snapshot(driver, "before_iframe_switch")
            iframe = wait.until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
            driver.switch_to.frame(iframe)
            wait.until(EC.element_to_be_clickable((By.XPATH, xpath))).click()
        time.sleep(2)
    log("Navigated to Event Show Wise New")


def get_iframe(driver: webdriver.Chrome, wait: WebDriverWait):
    driver.switch_to.default_content()
    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
    iframe = wait.until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
    driver.switch_to.frame(iframe)
    return iframe


def download_reports(driver: webdriver.Chrome, wait: WebDriverWait) -> None:
    time.sleep(10)
    get_iframe(driver, wait)

    for value in COMPANY_VALUES:
        log(f"Processing company code: {value}")

        get_iframe(driver, wait)

        # Set dropdown value via JS (avoids Select2 UI quirks)
        driver.execute_script("""
            var sel = document.getElementById('ddlCompanyCode');
            sel.value = arguments[0];
            sel.dispatchEvent(new Event('change'));
        """, value)

        wait.until(EC.element_to_be_clickable((By.ID, "btnShowReport"))).click()
        time.sleep(10)

        get_iframe(driver, wait)

        wait.until(EC.element_to_be_clickable((By.ID, "btnEXLExport"))).click()
        log(f"Exported report for {value}")
        time.sleep(5)


# ──────────────────────────────────────────────
# DB – insert downloaded files
# ──────────────────────────────────────────────

def process_files(processed_today: str) -> None:
    """Read all downloaded EventShowwiseRpt* Excel files and bulk-insert into MySQL."""

    files = [
        f for f in os.listdir(DOWNLOAD_FOLDER)
        if f.startswith("EventShowwiseRpt")
    ]

    if not files:
        log("No matching files found in downloads folder.")
        return

    # Single DB connection for all files
    conn   = mysql.connector.connect(**DB_CONFIG)
    cursor = conn.cursor()

    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for file in files:
        file_path = os.path.join(DOWNLOAD_FOLDER, file)
        log(f"Processing file: {file}")

        try:
            df = pd.read_excel(file_path)

            # Convert Excel serial dates in-place
            for col in ("Trans_Date", "Show_Date"):
                df[col] = df[col].apply(excel_serial_to_datetime)

            rows = [
                (
                    row["Bkg_Id"],
                    row["Trans_Date"],
                    row["Event_Name"],
                    row["Cinema_Name"],
                    row["Show_Date"],
                    row["Ticket_Qty"],
                    row["Seat_Info"],
                    "BMS",
                    current_time,   # insert_date
                    current_time,   # report_date
                    current_time,   # update_date
                    EVENT_TYPE,
                )
                for _, row in df.iterrows()
            ]

            if not rows:
                log(f"  No rows in {file}, skipping.")
                continue

            cursor.executemany(INSERT_QUERY, rows)
            conn.commit()
            log(f"  Inserted/updated {len(rows)} rows from {file}")

            # Move to processed/YYYY-MM-DD/
            os.makedirs(processed_today, exist_ok=True)
            dest = os.path.join(processed_today, file)
            log(f"  Moving to {dest}")
            shutil.move(file_path, dest)
            log(f"  Moved successfully")

        except Exception as exc:
            conn.rollback()
            log(f"  ERROR processing {file}: {exc}")

    cursor.close()
    conn.close()


# ──────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────

def main() -> None:
    os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)
    os.makedirs(PROCESSED_FOLDER, exist_ok=True)
    os.chdir(BASE_DIR)
    log("Script started")

    processed_today = make_processed_folder()

    driver = build_driver()
    wait   = WebDriverWait(driver, 30)

    try:
        login(driver, wait)
        navigate_to_report(driver, wait)
        download_reports(driver, wait)
    finally:
        driver.quit()
        log("Browser closed")

    process_files(processed_today)
    log("Script finished")


if __name__ == "__main__":
    main()
