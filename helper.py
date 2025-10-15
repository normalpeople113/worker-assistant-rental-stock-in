import os
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import tempfile

# Mapping lokasi ke XPath
LOCATION_XPATH_MAP = {
    "JAVAMIFI-RUKO": '//*[@id="view_1726_field_932_chzn_o_15"]',
    "Product Team": '//*[@id="view_1726_field_932_chzn_o_28"]',
    "Menara Caraka": '//*[@id="view_1726_field_932_chzn_o_27"]',
    "JAVAMIFI-RAWABOKOR": '//*[@id="view_1726_field_932_chzn_o_14"]',
    "JAVAMIFI-RAWABOKOR": '//*[@id="view_1726_field_932_chzn_o_14"]',
    "JAVAMIFI-BINTARO": '//*[@id="view_1726_field_932_chzn_o_7"]',
    "JAVAMIFI-BSD": '//*[@id="view_1726_field_932_chzn_o_8"]',
    "JAVAMIFI-AIRPORT YGY": '//*[@id="view_1726_field_932_chzn_o_2"]',
    "JAVAMIFI-SBY JUANDA": '//*[@id="view_1726_field_932_chzn_o_17"]',
    "JAVAMIFI-SBY WORKSHOP": '//*[@id="view_1726_field_932_chzn_o_19"]',
    "JAVAMIFI-MEDAN": '//*[@id="view_1726_field_932_chzn_o_13"]',
    "JAVAMIFI-BALI": '//*[@id="view_1726_field_932_chzn_o_3"]',
    
}

XPATHS = {
    "location_dropdown": '//*[@id="view_1726_field_932_chzn"]/a',
    "status_dropdown": '//select[@id="view_1726-field_961"][@name="field_961"]',
    "rental_input": '//input[@value="Select"]',
    "imei_input": '//input[@class="ui-autocomplete-input default"]',
    "imei_suggestion": '//*[@id="view_1726_field_1037_chzn_o_0"]',
    "submit_button": '//button[contains(text(), "Submit")]',
}

def get_location_xpath(name):
    xpath = LOCATION_XPATH_MAP.get(name)
    if not xpath:
        raise ValueError(f"❌ Location '{name}' tidak ditemukan di mapping.")
    print(f"[helper] ✅ Mapping location '{name}' → {xpath}")
    return xpath

def get_xpath(key):
    xpath = XPATHS.get(key)
    if not xpath:
        print(f"[helper] ❌ XPath key '{key}' tidak ditemukan di XPATHS")
    return xpath

def get_status_value(status):
    status_map = {
        "READY": 'READY',
        "BROKEN": 'BROKEN'
    }
    value = status_map.get(status.upper())
    if not value:
        raise ValueError(f"Unknown status: {status}")
    return value

def ensure_log_columns(df):
    """
    Pastikan kolom 'Logs' dan 'TimeStamp' ada dan bertipe string.
    Hindari FutureWarning saat assign string ke kolom kosong.
    """
    for col in ["Logs", "TimeStamp"]:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].astype(str)

def setup_driver(headless=True):
    chrome_options = Options()

    if headless:
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--log-level=3")
        chrome_options.add_argument("--disable-extensions")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")

    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option("useAutomationExtension", False)

    # 🔥 Folder unik biar gak bentrok
    user_data_dir = tempfile.mkdtemp()
    chrome_options.add_argument(f"--user-data-dir={user_data_dir}")
    print(f"[Driver] Using user-data-dir: {user_data_dir}")

    try:
        driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=chrome_options
        )
        driver.set_window_size(1920, 1080)
        return driver
    except Exception as e:
        print(f"[Driver Error] {type(e).__name__}: {e}")
        return None






