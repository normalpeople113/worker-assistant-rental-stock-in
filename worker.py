# worker.py

import os
import time
from datetime import datetime
from dotenv import load_dotenv
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import threading
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import StaleElementReferenceException
import tempfile
import shutil

from helper import (
    get_location_xpath,
    get_xpath,
    get_status_value,
    setup_driver,
)

exit_flag = False

# === CONFIG ===
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_DIR = os.path.join(SCRIPT_DIR, "config")
LOG_DIR = os.path.join(SCRIPT_DIR, "logs_worker")
SCREENSHOT_DIR = os.path.join(SCRIPT_DIR, "screenshots_worker")
os.makedirs(LOG_DIR, exist_ok=True)
os.makedirs(SCREENSHOT_DIR, exist_ok=True)
load_dotenv(dotenv_path=os.path.join(CONFIG_DIR, ".env"))

SHEET_URL = os.getenv("GSHEET_URL") or "https://docs.google.com/spreadsheets/d/1mzepZm0EsqUIppLeIMZxRRSROzDMCnsvrEN5RYGJvL4/edit?gid=0#gid=0"
JSON_CRED = os.getenv("GSHEET_JSON") or r"/www/wwwroot/Worker_Rental_Stock_In/active-bolt-398921-fdef5f0fc06a.json"
WORKER_SHEETS = [f"Worker-{i}" for i in range(1, 6)]  # Worker-1 sampai Worker-5

def log(msg):
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(f"[{timestamp}] {msg}")
    with open(os.path.join(LOG_DIR, "log.txt"), "a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] {msg}\n")

def save_step(driver, name):
    path = os.path.join(SCREENSHOT_DIR, f"{name}.png")
    driver.save_screenshot(path)

def wait_visible_xpath(driver, xpath, timeout=15):
    try:
        return WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((By.XPATH, xpath)))
    except Exception as e:
        log(f"⚠️ wait_visible_xpath gagal: {xpath} → {type(e).__name__}")
        save_step(driver, f"wait_visible_failed_{xpath[-10:].replace('/', '_')}")
        return None

def force_input(driver, elem, value, label="Input"):
    try:
        elem.click()
        elem.clear()
        for digit in str(value):
            elem.send_keys(digit)
        for evt in ["input", "change", "blur"]:
            driver.execute_script(f"arguments[0].dispatchEvent(new Event('{evt}', {{ bubbles: true }}));", elem)
        log(f"⌨️ {label} diketik: {value}")
    except Exception as e:
        log(f"❌ Gagal input {label}: {type(e).__name__} - {e}")
        save_step(driver, f"error_input_{label.lower()}")

def retry_action(action_func, max_retry=3, *args, **kwargs):
    for attempt in range(1, max_retry + 1):
        try:
            return action_func(*args, **kwargs)
        except StaleElementReferenceException as e:
            log(f"⚠️ Percobaan ke-{attempt}: StaleElementReferenceException, retry ...")
            time.sleep(1)
        except Exception as e:
            raise
    raise Exception(f"Gagal setelah {max_retry} percobaan: {action_func.__name__}")

def select_chosen_option(driver, dropdown_xpath, option_xpath, label):
    for attempt in range(3):
        try:
            dropdown = wait_visible_xpath(driver, dropdown_xpath)
            if dropdown is None:
                log(f"⚠️ Dropdown {label} tidak ditemukan (attempt {attempt+1})")
                time.sleep(1)
                continue
            dropdown.click()
            time.sleep(0.5)

            option = wait_visible_xpath(driver, option_xpath)
            if option is None:
                log(f"⚠️ Option {label} tidak ditemukan (attempt {attempt+1})")
                time.sleep(1)
                continue
            option.click()
            log(f"✅ {label} dipilih")
            save_step(driver, f"step_chosen_{label.lower()}")
            return True
        except StaleElementReferenceException:
            log(f"⚠️ StaleElementReferenceException pada {label}, retry ...")
            time.sleep(1)
        except Exception as e:
            log(f"❌ Gagal pilih {label}: {type(e).__name__} - {e}")
            save_step(driver, f"error_chosen_{label.lower()}")
            time.sleep(1)
    return False

def select_dropdown_by_value(driver, xpath, value, label="Dropdown"):
    try:
        elem = wait_visible_xpath(driver, xpath)
        if elem:
            Select(elem).select_by_value(value)
            log(f"✅ {label} dipilih: {value}")
            save_step(driver, f"step_dropdown_{label.lower()}")
            return True
        else:
            log(f"❌ {label} tidak ditemukan: {xpath}")
            return False
    except Exception as e:
        log(f"❌ Gagal pilih {label}: {type(e).__name__} - {e}")
        save_step(driver, f"error_dropdown_{label.lower()}")
        return False

def click_suggestion(driver, xpath):
    try:
        suggestion = wait_visible_xpath(driver, xpath)
        suggestion.click()
        log("✅ IMEI diklik")
        save_step(driver, "step_imei_suggestion")
        return True
    except Exception as e:
        log(f"❌ Gagal klik IMEI: {type(e).__name__} - {e}")
        save_step(driver, "error_imei_suggestion")
        return False

def login_knack(driver, url, email, password):
    driver.get(url)
    log("🔄 Buka halaman login Knack")
    try:
        wait_visible_xpath(driver, '//input[@type="email"]').send_keys(email)
        driver.find_element(By.XPATH, '//input[@type="password"]').send_keys(password)
        log("📧 Email dan password diketik")
        for xpath in [
            '//button[contains(text(), "Login")]',
            '//button[contains(text(), "Sign In")]',
            '//input[@type="submit"]',
            '//*[@id="submit"]'
        ]:
            try:
                driver.find_element(By.XPATH, xpath).click()
                log(f"✅ Klik login berhasil: {xpath}")
                break
            except:
                # Hapus log XPath gagal biar nggak spam log
                pass
        if wait_visible_xpath(driver, '//div[contains(@id, "view_") and contains(@class, "kn-view")]'):
            log("✅ Login sukses, view Knack siap")
            return True
        else:
            log("❌ View Knack tidak muncul setelah login")
            save_step(driver, "error_view_knack")
            return False
    except Exception as e:
        log(f"❌ Gagal login otomatis: {type(e).__name__} - {e}")
        save_step(driver, "error_login")
        return False

def submit_row(driver, row):
    location = str(row["Location"])
    imei = str(row["IMEI"])
    status = str(row["Status"])

    xpath_location_dropdown = get_xpath("location_dropdown")
    xpath_location_option = get_location_xpath(location)
    xpath_status = get_xpath("status_dropdown")
    xpath_rental_input = get_xpath("rental_input")
    xpath_imei_suggestion = get_xpath("imei_suggestion")
    xpath_submit = get_xpath("submit_button")

    if not xpath_location_option:
        raise Exception(f"Gagal pilih Location: {location} (mapping tidak ditemukan)")

    if not retry_action(select_chosen_option, 3, driver, xpath_location_dropdown, xpath_location_option, f"Location: {location}"):
        raise Exception(f"Gagal pilih Location: {location}")

    status_value = get_status_value(status)
    if not retry_action(select_dropdown_by_value, 3, driver, xpath_status, status_value, f"Status: {status}"):
        raise Exception(f"Gagal pilih Status: {status}")

    trigger_elem = wait_visible_xpath(driver, xpath_rental_input)
    if trigger_elem:
        trigger_elem.click()
        log("🖱️ Trigger IMEI diklik")
        save_step(driver, "step_trigger_imei")
    else:
        raise Exception("Gagal menemukan input IMEI")

    input_elem = driver.execute_script("return document.activeElement")
    if input_elem:
        force_input(driver, input_elem, imei, label="IMEI")
        save_step(driver, "step_imei_input_active")
        actions = ActionChains(driver)
        actions.send_keys(Keys.F12).perform()
        log("🎹 F12 dikirim setelah input IMEI untuk trigger dropdown")
        time.sleep(1)
    else:
        raise Exception("Gagal ambil activeElement untuk input IMEI")

    if not retry_action(click_suggestion, 3, driver, xpath_imei_suggestion):
        raise Exception("Gagal klik Suggestion IMEI")

    try:
        driver.execute_script("""const active = document.querySelector('.active-result'); if (active) active.click();""")
        time.sleep(0.5)
        log("🧹 Dropdown aktif ditutup")
    except:
        log("⚠️ Tidak ada dropdown aktif yang perlu ditutup")

    submit_elem = wait_visible_xpath(driver, xpath_submit)
    if submit_elem:
        driver.execute_script("arguments[0].scrollIntoView(true);", submit_elem)
        driver.execute_script("arguments[0].click();", submit_elem)
        log("🚀 Klik tombol Submit (force click)")
        save_step(driver, "step_submit_force")
    else:
        raise Exception("Gagal menemukan tombol Submit")

    log("🎉 Submit selesai, workflow sukses")
    return True

def get_gsheet_client(json_credential_path):
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(json_credential_path, scope)
    client = gspread.authorize(creds)
    return client

def get_or_create_sheet(spreadsheet, sheet_name):
    try:
        return spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        return spreadsheet.add_worksheet(title=sheet_name, rows="1000", cols="20")

def ensure_log_columns(sheet):
    headers = sheet.row_values(1)
    updated = False
    if "Logs" not in headers:
        sheet.update_cell(1, len(headers)+1, "Logs")
        headers.append("Logs")
        updated = True
    if "TimeStamp" not in headers:
        sheet.update_cell(1, len(headers)+1, "TimeStamp")
        headers.append("TimeStamp")
        updated = True
    return headers

def get_knack_account(sheet):
    # Email di kolom I1 (kolom ke-9), Password di kolom K1 (kolom ke-11)
    email = sheet.cell(1, 9).value.strip() if sheet.cell(1, 9).value else ""
    password = sheet.cell(1, 11).value.strip() if sheet.cell(1, 11).value else ""
    return email, password

def process_all_workers(sheet_url, json_credential_path, worker_sheets):
    client = get_gsheet_client(json_credential_path)
    spreadsheet = client.open_by_url(sheet_url)

    driver = setup_driver(headless=True)
    if not driver:
        log("Driver gagal di-load, workflow dihentikan.")
        return

    def input_thread():
        global exit_flag
        while True:
            user_input = input()
            if user_input.strip().lower() == 'q':
                exit_flag = True
                break

    t = threading.Thread(target=input_thread, daemon=True)
    t.start()

    try:
        while not exit_flag:
            for sheet_name in worker_sheets:
                sheet = get_or_create_sheet(spreadsheet, sheet_name)
                headers = ensure_log_columns(sheet)
                logs_col = headers.index("Logs") + 1
                ts_col = headers.index("TimeStamp") + 1

                # Ambil email & password dari sheet
                try:
                    email, password = get_knack_account(sheet)
                except Exception as e:
                    log(f"[{sheet_name}] Gagal ambil akun: {e}")
                    continue

                # Login Knack dengan akun worker ini
                url = os.getenv("FORM_URL")
                if not login_knack(driver, url, email, password):
                    log(f"[{sheet_name}] Login gagal, skip sheet ini")
                    continue
                log(f"{email} telah login di {sheet_name}")  # <-- Perbaikan log di sini

                data = sheet.get_all_records()
                for idx, row in enumerate(data, start=2):
                    imei = str(row.get("IMEI", "")).strip()
                    if not imei or str(row.get("Logs", "")).startswith("✅"):
                        continue

                    log_msg = ""
                    try:
                        submit_row(driver, row)
                        log_msg = "✅ Submit sukses"
                    except Exception as e:
                        log_msg = f"❌ Error: {type(e).__name__} - {e}"

                    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    sheet.update_cell(idx, logs_col, log_msg)
                    sheet.update_cell(idx, ts_col, timestamp)
                    log(f"[{sheet_name}] Row {idx-1} updated: {log_msg} at {timestamp}")

            log("⏳ Menunggu data baru di semua worker sheets... (atau ketik 'q'+Enter untuk keluar)")
            time.sleep(10)
    finally:
        driver.quit()
        log("🛑 Tutup browser")

def worker_process(sheet_url, json_credential_path, sheet_name):
    client = get_gsheet_client(json_credential_path)
    spreadsheet = client.open_by_url(sheet_url)
    sheet = get_or_create_sheet(spreadsheet, sheet_name)
    headers = ensure_log_columns(sheet)
    logs_col = headers.index("Logs") + 1
    ts_col = headers.index("TimeStamp") + 1

    try:
        email, password = get_knack_account(sheet)
        # Ganti log, jangan tampilkan password
        log(f"[{sheet_name}] Email: '{email}'")
    except Exception as e:
        log(f"[{sheet_name}] Gagal ambil akun: {e}")
        return

    if not email or not password:
        log(f"[{sheet_name}] Akun kosong, worker tidak dijalankan.")
        return

    url = os.getenv("FORM_URL")
    with tempfile.TemporaryDirectory(prefix=f"profile_{sheet_name}_") as profile_dir:
        try:
            driver = setup_driver(headless=True, profile_dir=profile_dir)
        except Exception as e:
            log(f"[{sheet_name}] ERROR Chrome gagal dibuka: {e}")
            return

        try:
            if not login_knack(driver, url, email, password):
                log(f"[{sheet_name}] Login gagal, skip sheet ini")
                return
            log(f"{email} telah berhasil login knack di {sheet_name}")
            time.sleep(2)  # Jeda setelah login

            while True:
                data = sheet.get_all_records()
                for idx, row in enumerate(data, start=2):
                    imei = str(row.get("IMEI", "")).strip()
                    if not imei or str(row.get("Logs", "")).startswith("✅"):
                        continue

                    log_msg = ""
                    try:
                        submit_row(driver, row)
                        log_msg = "✅ Submit sukses"
                        time.sleep(2)  # Jeda setelah submit
                    except Exception as e:
                        log_msg = f"❌ Error: {type(e).__name__} - {e}"
                        time.sleep(2)  # Jeda setelah error

                    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    # Ganti dua update_cell jadi satu update
                    sheet.update(f"{chr(64+logs_col)}{idx}:{chr(64+ts_col)}{idx}", [[log_msg, timestamp]])
                    time.sleep(2)  # Jeda setelah update
                    log(f"[{sheet_name}] Row {idx-1} updated: {log_msg} at {timestamp}")

                time.sleep(10)
        finally:
            driver.quit()
            log(f"🛑 Tutup browser untuk {sheet_name}")

def setup_driver(headless=True, profile_dir=None):
    chrome_options = Options()
    if headless:
        chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    if profile_dir:
        chrome_options.add_argument(f"--user-data-dir={profile_dir}")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    return driver

if __name__ == "__main__":
    threads = []
    for sheet_name in WORKER_SHEETS:
        t = threading.Thread(target=worker_process, args=(SHEET_URL, JSON_CRED, sheet_name), daemon=True)
        t.start()
        threads.append(t)
    # Tunggu semua thread selesai (atau tekan Ctrl+C untuk keluar)
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        log("🛑 Semua worker dihentikan oleh user")
