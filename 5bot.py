import requests
import json
import time
import logging
import os
import sys
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select

from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException
)
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import threading
from queue import Queue
from threading import Lock

class ThreadRoutingHandler(logging.Handler):
    def __init__(self, log_dir):
        super().__init__()
        self.log_dir = log_dir
        self.handlers = {}
        self.formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

    def emit(self, record):
        thread_name = record.threadName
        if thread_name.startswith("Bot_"):
            if thread_name not in self.handlers:
                os.makedirs(self.log_dir, exist_ok=True)
                log_file = os.path.join(self.log_dir, f"{thread_name}.log")
                h = logging.FileHandler(log_file)
                h.setFormatter(self.formatter)
                self.handlers[thread_name] = h
            self.handlers[thread_name].emit(record)

# Initial global logging setup
log_dir = "logs"
os.makedirs(log_dir, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("happy_flow.log"),
        logging.StreamHandler(),
        ThreadRoutingHandler(log_dir)
    ]
)

excel_lock = Lock()
EXCEL_FILE = "order_report.xlsx"

def log_to_excel(order_id, status, error_msg=""):
    with excel_lock:
        try:
            import pandas as pd
            # Create a clean DataFrame if file doesn't exist
            if os.path.exists(EXCEL_FILE):
                df = pd.read_excel(EXCEL_FILE)
            else:
                df = pd.DataFrame(columns=["Order Number", "Status", "Error Details", "Timestamp"])
            
            new_row = {
                "Order Number": str(order_id),
                "Status": status,
                "Error Details": str(error_msg),
                "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            # Append new record
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            df.to_excel(EXCEL_FILE, index=False)
        except Exception as e:
            logging.error(f"Failed to log to Excel: {e}")

# =====================================================
# BALANCE & HISTORY HELPERS
# =====================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
BALANCE_FILE = os.path.join(BASE_DIR, "balance.txt")
ORDERS_COUNT_FILE = os.path.join(BASE_DIR, "orders_count.txt")
HISTORY_FILE = os.path.join(BASE_DIR, "processed_orders.json")
COST_PER_ORDER = 0.75

def get_balance():
    if not os.path.exists(BALANCE_FILE): return 0.0
    try:
        with open(BALANCE_FILE, "r") as f:
            content = f.read().strip()
            return float(content) if content else 0.0
    except:
        return 0.0

def update_balance(val):
    with open(BALANCE_FILE, "w") as f:
        f.write(str(round(val, 2)))

def update_history(entry):
    history = []
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "r") as f:
                history = json.load(f)
        except:
            pass
    
    entry["timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    history.append(entry)
    
    # Save last 1000 records
    with open(HISTORY_FILE, "w") as f:
        json.dump(history[-1000:], f, indent=4)
    
    # Update orders count
    processed_count = len([e for e in history if e.get("status") == "Processed"])
    with open(ORDERS_COUNT_FILE, "w") as f:
        f.write(str(processed_count))

# =====================================================
# API CONFIG
# =====================================================

API_URL = "https://abadmin.agribegri.com/API/api_get_confirm_order_details.php"

API_HEADERS = {
    "Authorization": "MiwiaXNzIjoiaHR0cHM6Ly9hcGl2Mi5zaGlwcm9ja2V0LmluL3YxL2V4dGVybmFsL2F1dGgvbG9naW4iLCJpYXQiOjE2NjUwNTkzNDAsImV4cCI6MTY2NTkyMzM0MCwibmJmIjoxNjY1MDU5MzQwLCJqdGkiOiI5OTNlTzkzRndMOGlIY1dBIn0.6DaySnpwyBM0XInHywTgxlddCFLgx_WxzRD4jKSheq",
    "Cookie": "ab_temp_order_id=75058183",
    "User-Agent": "Mozilla/5.0"
}

API_PAYLOAD = {
    "order_id": "75058183"
}

# =====================================================
# ORDER FILTER LIMITS
# =====================================================

WEIGHT_LIMIT_GRAM = 7000
AMOUNT_LIMIT = 8000

# =====================================================
# LOGIN CONFIG
# =====================================================

if len(sys.argv) >= 5:
    USERNAME = sys.argv[1]
    PASSWORD = sys.argv[2]
    OTP = sys.argv[3]
    TARGET_ORDER_ID = sys.argv[4]
elif len(sys.argv) >= 4:
    USERNAME = sys.argv[1]
    PASSWORD = sys.argv[2]
    OTP = sys.argv[3]
    TARGET_ORDER_ID = None
else:
    USERNAME = "Namrata"
    PASSWORD = "Namrata@2026"
    OTP = "123456"
    TARGET_ORDER_ID = None

SENDER_EMAIL = "dispatch.agribegri@gmail.com"
SENDER_PASSWORD = "ndka txuk ftho pwyn"


# =====================================================
# PATHS
# =====================================================

download_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "downloads")
if not os.path.exists(download_path):
    os.makedirs(download_path)

# =====================================================
# CHROME DRIVER

# =====================================================

order_queue = Queue()

processed_orders = set()

processed_lock = Lock()

def send_email(to_email, subject, body, attachments):
    try:
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        for file_path in attachments:
            if os.path.exists(file_path):
                filename = os.path.basename(file_path)
                with open(file_path, "rb") as attachment:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header("Content-Disposition", f"attachment; filename= {filename}")
                    msg.attach(part)

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        text = msg.as_string()
        server.sendmail(SENDER_EMAIL, to_email, text)
        server.quit()
        logging.info(f"Email sent successfully to {to_email}")
        return True
    except Exception as e:
        logging.error(f"Failed to send email: {e}")
        return False


def init_driver(bot_id):
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    
    # Hide automation flags
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    
    # Force absolute path for downloads
    dl_path = os.path.abspath(os.path.join(os.getcwd(), f"downloads/bot_{bot_id}"))
    os.makedirs(dl_path, exist_ok=True)

    # Bypass security blocks and mixed content issues
    options.add_argument("--ignore-certificate-errors")
    options.add_argument("--allow-running-insecure-content")
    options.add_argument("--safebrowsing-disable-download-protection")
    options.add_argument("--disable-features=SafeBrowsing")
    
    prefs = {
        "download.default_directory": dl_path,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": False,
        "safebrowsing.disable_download_protection": True,
        "safebrowsing_for_trusted_sources_enabled": False,
        "profile.default_content_setting_values.automatic_downloads": 1,
        "profile.default_content_settings.popups": 0,
        "plugins.always_open_pdf_externally": True
    }
    options.add_experimental_option("prefs", prefs)



    driver = webdriver.Chrome(options=options)
    
    # Hide webdriver property
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": """
            Object.defineProperty(navigator, 'webdriver', {
                get: () => undefined
            })
        """
    })
    
    wait = WebDriverWait(driver, 20)
    return driver, wait, dl_path
# =====================================================
# LOGIN FUNCTION
# =====================================================

def login(driver, wait):

    driver.get("https://abadmin.agribegri.com/admin/")

    logging.info("Opening Login Page...")

    # =================================================
    # USERNAME
    # =================================================

    username_input = wait.until(
        EC.presence_of_element_located((By.ID, "username"))
    )

    username_input.clear()
    username_input.send_keys(USERNAME)

    logging.info("Username Entered")

    driver.find_element(By.ID, "btnSubmit").click()

    time.sleep(2)

    # =================================================
    # USERNAME ERROR
    # =================================================

    try:
        username_error = driver.find_element(
            By.CSS_SELECTOR,
            ".login-alert-username .alert-error"
        )

        if username_error.is_displayed():
            logging.error(f"ERROR: {username_error.text}")
            driver.quit()
            exit()

    except:
        pass

    # =================================================
    # PASSWORD
    # =================================================

    password_input = wait.until(
        EC.visibility_of_element_located((By.ID, "password"))
    )

    password_input.clear()
    password_input.send_keys(PASSWORD)

    logging.info("Password Entered")

    driver.find_element(By.ID, "btnSubmit").click()

    time.sleep(2)

    # =================================================
    # PASSWORD ERROR
    # =================================================

    try:
        password_error = driver.find_element(
            By.CSS_SELECTOR,
            ".login-alert-password .alert-error"
        )

        if password_error.is_displayed():
            logging.error(f"ERROR: {password_error.text}")
            driver.quit()
            exit()

    except:
        pass

    # =================================================
    # OTP
    # =================================================

    otp_input = wait.until(
        EC.visibility_of_element_located((By.ID, "otp"))
    )

    otp_input.clear()
    otp_input.send_keys(OTP)

    logging.info("OTP Entered")

    driver.find_element(By.ID, "btnSubmit").click()

    time.sleep(5)

    # =================================================
    # OTP ERROR
    # =================================================

    try:
        otp_error = driver.find_element(
            By.CSS_SELECTOR,
            ".login-alert-otp .alert-error"
        )

        if otp_error.is_displayed():
            logging.error(f"ERROR: {otp_error.text}")
            driver.quit()
            exit()

    except:
        pass

    logging.info("LOGIN SUCCESSFUL")
    logging.info("Dashboard Opened")


# =====================================================
# SHIPPING THROUGH COURIER FILTER
# =====================================================

def select_shipping_through_courier(driver, wait):



    logging.info(
        "Locating Shipping Through dropdown..."
    )

    # ============================================
    # OPEN SHIPPING DROPDOWN
    # ============================================

    shipping_dropdown_btn = wait.until(
        EC.presence_of_element_located((
            By.XPATH,
            "//button[contains(@class,'multiselect') "
            "and contains(@title, 'Shipping Through')]"
        ))
    )

    driver.execute_script(
        "arguments[0].click();",
        shipping_dropdown_btn
    )

    logging.info(
        "Shipping Through dropdown opened"
    )

    time.sleep(1)

    # ============================================
    # SELECT SHIPPING THROUGH COURIER
    # ============================================

    courier_checkbox = wait.until(
        EC.presence_of_element_located((
            By.XPATH,
            "//label[contains(., "
            "'Shipping Through Courier')]//input"
        ))
    )

    driver.execute_script(
        "arguments[0].click();",
        courier_checkbox
    )

    logging.info(
        "Shipping Through Courier selected"
    )

    time.sleep(1)

    # ============================================
    # CLOSE DROPDOWN
    # ============================================

    driver.execute_script(
        "arguments[0].click();",
        shipping_dropdown_btn
    )

    logging.info(
        "Shipping dropdown closed"
    )

    time.sleep(1)

# =====================================================
# APPLY FILTER FUNCTION
# =====================================================

def apply_filter(driver, wait, target_order_id=None):

    logging.info("Opening Manage Orders page")

    driver.get(
        'https://abadmin.agribegri.com/admin/manage_orders.php'
    )

    time.sleep(3)

    wait = WebDriverWait(driver, 30)

    logging.info("Manage Orders page loaded")

    # ================= DATE FILTER =================

    today_date = datetime.now()

    from_date_value = (
        today_date - timedelta(days=30)
    ).strftime("%Y-%m-%d")

    to_date_value = today_date.strftime("%Y-%m-%d")

    logging.info(
        f"Applying Date Filter | "
        f"From: {from_date_value} | "
        f"To: {to_date_value}"
    )

    # FROM DATE
    driver.execute_script("""
        let fromInput = document.getElementById('from_date');

        fromInput.removeAttribute('readonly');

        fromInput.value = arguments[0];

        fromInput.dispatchEvent(
            new Event('change', { bubbles: true })
        );
    """, from_date_value)

    time.sleep(1)

    # TO DATE
    driver.execute_script("""
        let toInput = document.getElementById('to_date');

        toInput.removeAttribute('readonly');

        toInput.value = arguments[0];

        toInput.dispatchEvent(
            new Event('change', { bubbles: true })
        );
    """, to_date_value)

    time.sleep(1)

    logging.info("Date Filter Applied")

    # ================= ORDER STATUS =================

    logging.info("Selecting Confirm Status")

    status_dropdown_btn = wait.until(
        EC.element_to_be_clickable((
            By.XPATH,
            "//button[contains(@title, 'Order Status')]"
        ))
    )

    driver.execute_script(
        "arguments[0].click();",
        status_dropdown_btn
    )

    time.sleep(1)

    confirm_option = wait.until(
        EC.presence_of_element_located((
            By.XPATH,
            "//input[@value='Confirm']"
        ))
    )

    driver.execute_script(
        "arguments[0].click();",
        confirm_option
    )

    time.sleep(1)

    # CLOSE DROPDOWN
    driver.execute_script(
        "arguments[0].click();",
        status_dropdown_btn
    )

    logging.info("Confirm Status Selected")

    
    
    # ================================================
    # SHIPPING THROUGH COURIER FILTER
    # ================================================

    select_shipping_through_courier(driver, wait)

    # ================= SEARCH BUTTON =================

    search_button = driver.find_element(
        By.ID,
        'srchSubmit'
    )

    driver.execute_script(
        "arguments[0].click();",
        search_button
    )

    logging.info("Search Button Clicked")

    # WAIT TABLE
    wait.until(
        EC.presence_of_element_located((
            By.CSS_SELECTOR,
            "table tbody tr"
        ))
    )

    logging.info("Orders Table Loaded")


    # ================================================
    # SELECT 100 ENTRIES
    # ================================================

    try:

        logging.info("Selecting 100 entries")

        entries_dropdown = wait.until(
            EC.presence_of_element_located((
                By.NAME,
                "dyntable_length"
            ))
        )

        driver.execute_script(
            """
            arguments[0].value = '100';
            arguments[0].dispatchEvent(
                new Event('change')
            );
            """,
            entries_dropdown
        )

        time.sleep(3)

        logging.info("100 entries selected")

    except Exception as e:

        logging.warning(
            f"100 entries dropdown failed: {e}"
        )

    # ================================================
    # CLICK LAST PAGE IF PRESENT
    # ================================================

    try:

        logging.info(
            "Checking for Last button..."
        )

        time.sleep(3)

        # WAIT FOR PAGINATION AREA
        wait.until(
            EC.presence_of_element_located((
                By.CLASS_NAME,
                "dataTables_paginate"
            ))
        )

        # FIND LAST BUTTON
        last_buttons = driver.find_elements(
            By.XPATH,
            "//a[contains(@class,'cus_page_act') and contains(text(),'Last')]"
        )

        if last_buttons:

            last_button = last_buttons[0]

            logging.info(
                "Last button found"
            )

            # SCROLL TO BUTTON
            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});",
                last_button
            )

            time.sleep(1)

            # CLICK USING JS
            driver.execute_script(
                "arguments[0].click();",
                last_button
            )

            logging.info(
                "Last button clicked"
            )

            # WAIT TABLE RELOAD
            time.sleep(5)

            logging.info(
                "Moved to last page"
            )

        else:

            logging.info(
                "Last button not present"
            )

    except Exception as e:

        logging.warning(
            f"Last page click failed: {e}"
        )


    # ================================================
    # SEARCH FIRST ORDER
    # ================================================

    if target_order_id:

        try:

            logging.info(
                f"Searching Order: {target_order_id}"
            )

            search_box = wait.until(
                EC.element_to_be_clickable((
                    By.XPATH,
                    "//input[@aria-controls='dyntable']"
                ))
            )

            driver.execute_script("arguments[0].value = '';", search_box)
            search_box.send_keys(target_order_id)
            search_box.send_keys(u'\ue007') # Press Enter

            time.sleep(3)

            logging.info(
                f"Order searched: {target_order_id}"
            )

        except Exception as e:

            logging.warning(
                f"Search failed: {e}"
            )
# =====================================================
# API FUNCTION
# =====================================================

def fetch_orders():

    logging.info("Fetching Orders From API...")

    response = requests.post(
        API_URL,
        headers=API_HEADERS,
        data=API_PAYLOAD
    )

    logging.info(f"API Status: {response.status_code}")

    try:

        data = response.json()

        # Save Full JSON
        with open("all_orders.json", "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4, ensure_ascii=False)

        logging.info("JSON Saved")

        return data

    except Exception as e:

        logging.error(f"JSON Error: {e}")
        logging.error(response.text)

        return None


# =====================================================
# SELLER PICKUP MAP
# =====================================================

SELLER_PICKUP_MAP = { "abhihsek of be bearings": "OEM India Pvt. Ltd.", "abhijit buch": "Abhijit Buch", "abhilash nair": "Sonikraft", "abhinandan traders": "Abhinandan Traders", "abhishek": "Akshar Farmtech LLP", "abhishek jhawar": "ABHISHEK JHAWAR", "aipm": "AIPM", "airdrops irrigation pvt ltd": "Airdrops Irrigation Pvt Ltd", "aj kisan agrotech": "AJ KISAN AGROTECH", "akash dhanokar": "Apras Polymers & Engineering Co Pvt Ltd", "akhil garg of rs enterprises": "Akhil Garg OF SNM TELECARE", "akshay joshi": "Akshay Joshi", "allbata agriculture biotech pvt ltd": "Albata", "aman agencies": "AMAN AGENCIES", "amit singla": "Amit singla", "amruth organic fertilizers": "Amruth Organic Fertilizers", "amuthalakshmi agro organic,": "Amuthalakshmi Agro Organic,", "anand agro care": "Anand Agro1", "anil packaging": "Anil Packaging", "anjali pitre of urvara": "Urvara Marketing Solutions Pvt Ltd", "annadata organic": "Annadata Organic", "ansh chhabra": "Ansh Chhabra", "ashish tiwari": "RamK Agrotech", "ayushi bansal": "BHUMI AGRO INDUSTRIES", "aziz ahamed": "Sunjaree Fertilizers", "berrysun agro science pvt. ltd.": "Berrysun Agro Science Pvt. Ltd.", "bharat parmar": "Shyam innovations", "bharath r": "Microbi Agrotech Pvt. Ltd.", "bhavani seeds center": "BHAVANI SEEDS CENTER", "bhumi polymers private limited": "BHUMI POLYMERS PRIVATE LIMITED", "biosun agri crop science pvt. ltd": "Biosun Agri Crop Science Pvt. Ltd.", "bm bio energy": "BM BIO ENERGY", "brijesh suresh somani": "Kamal Agrotech", "chetna patidar": "Rashail Agro", "chirag sojitra of bacf": "Chirag Sojitra", "cubic fertichem pvt ltd": "Cubic Fertichem Pvt Ltd", "deepak sharma": "Deepak Sharma", "dhaval gadhiya": "Gujarat Agri-Chem Industries Private Limited", "dr enterprises": "DR ENTERPRISES", "durga beej bhandar": "DURGA BEEJ BHANDAR", "earth innovation": "EARTH INNOVATION", "easykrishi private limited": "Easykrishi Private Limited", "ecotika india": "1 Ecotika India", "essential biosciences": "Essential Biosciences", "excellar production private limited": "EXCELLAR PRODUCTION PRIVATE LIMITED", "exosolar private limited": "Exosolar Private Limited", "farm chem india private limited": "FARM CHEM INDIA PRIVATE LIMITED", "farmson biotech pvt ltd": "FARMSON BIOTECH PVT LTD", "fenton technologies": "Fenton Technologies", "gabani brothers limited": "GABANI BROTHERS LIMITED", "gaiagen technologies private limited": "Gaiagen Technologies Private Limited", "gassin pierre pvt ltd": "Gassin Pierre Pvt Ltd", "gaurav kakkar": "GAURAV KAKKAR", "geolife agritech india pvt. ltd": "Geolife Agritech India Pvt. Ltd", "greatindos": "Greatindos", "green raksha agro seva": "Green Raksha Agro Seva", "green revolution": "Green Revolution", "greeno biotech": "Greeno Biotech", "greenovate agrotech pvt ltd": "1Greenovate Agrotech Pvt Ltd", "greenpeace agro industries": "1Greenpeace Agro", "greenvayu innovations private limited": "GREENVAYU INNOVATIONS PRIVATE LIMITED", "gumtree traps pvt.ltd.": "Gumtree Traps", "gupta agro enterprises": "GUPTA AGRO ENTERPRISES", "gvd electricals": "GVD ELECTRICALS", "harmony ecotech pvt ltd": "Harmony Ecotech Pvt Ltd", "hemendra patidar": "Hemendra Patidar", "hifield-ag chem india pvt ltd.": "Hifield-AG", "hussain lokhandwala": "HM Organics", "infinite biotech co.": "Infinite Biotech Co", "ishan mistry of sk inter": "ISHAN MISTRY Of SK Inter", "ishan sikka": "TUFFPAULIN", "jagannath biotech pvt ltd": "Jagannath Bio Tech Pvt Ltd", "janatha fish meal and oil products": "JANATHA FISH MEAL AND OIL PRODUCTS", "jignesh parmar": "Singhal Industries Pvt. Ltd.", "jinendra magdum": "Jinendra Magdum", "jivit seeds pvt. ltd.": "Jivit Seeds Pvt. Ltd.", "jiya enterprise": "JIYA ENTERPRISE", "kalpana tomar of haridwar.shoppee": "Kalpana Tomar Of Haridwar.Shoppee", "kamna dhawan": "kamna Dhawan", "kanchan kushwaha": "Kanchan Kushwaha", "kandasamy agency": "Kandasamy Agency", "kap associates": "KAP Associates", "kartavya agritech private limited": "KARTAVYA AGRITECH PRIVATE LIMITED", "kashinath c chilshetty": "Kashinath C Chilshetty", "katra fertilizers and chemicals private limited": "Katra", "katyayani organics": "Katyayani Organics", "keshav goyal": "Keshav Goyal", "kisan agrotech": "KISAN AGROTECH", "kodagu agritec private limited": "Kodagu Agritec Private Limited", "krishna seeds farm": "Goldi", "krushna krushi": "KRUSHNA KRUSHI", "kukkar spary centre": "Kukkar Spary Centre", "lta trust": "LTA Trust", "maltose bio innovations private limited": "MALTOSE BIO", "manan patel": "VINSPIRE AGROTECH (I) PVT LTD", "manik": "Manik", "martanbhai patel": "M. N. AGRO INDUSTRIES", "megatex protective fabrics private limited": "MEGATEX PROTECTIVE FABRICS PRIVATE LIMITED", "mettur agro traders": "METTUR AGRO TRADERS", "mipatex india": "Mipatex", "mitrasena": "Mitrasena", "mohan merchandise pvt. ltd": "Mohan Merchandise Private Limited", "monika malakar": "Modish Tractoraurkisan Pvt Ltd", "mrs. mitali hingorani": "Mrs. Mitali Hingorani", "muhammed meerasha": "Muhammed meerasha", "nathsagar bio genetics private limited": "Nathsagar Bio-Genetics Pvt Ltd", "navik organic products": "Navik Organic Products", "navin panchal": "Navin Panchal1", "nihar jain of siesto, chhattisgarh.": "SRT Agro Science Pvt. Ltd.", "nilesh deshpande, maharashtra.": "Urvara Marketing Solutions Pvt Ltd", "nimesh patel": "Nimesh Patel", "ninganagouda biradar": "Ninganagouda Biradar", "octopus crop care": "Octopus Crop Care", "ojas bhattad": "Ojas bhattad", "parth savaliya": "EUREKA SEEDS INDIA PRIVATE LIMITED", "pawan kumar": "Pawan kumar", "pheromone chemicals": "Pheromone Chemicals", "pioneer agro industry": "Pioneer agro Industry", "piyush garg": "Piyush", "piyush kataria": "PIYUSH KATARIA", "prabhat krishi kendra": "Prabhat Krishi Kendra", "pratiksha": "Ajay Bio-Tech(India)Ltd.", "radhe agri center": "Radhe Agri Center", "rahul garg": "Rahul Garg", "rahul jain": "Aquagri Processing Private Ltd", "raj kumar jaiswal": "DHANDA AGRO CHEMICAL INDUSTRIES", "rajib sarkar": "Rajib Sarkar", "rakhi naskar": "Rakhi Naskar", "ramandeep singh": "Ramandeep singh", "ravikumar patel": "Jakson Seeds Private Limited", "ravindra pofalkar": "Ravindra Pofalkar", "risepect enterprise": "Risepect Enterprise", "rk chemicals": "dobariya sunil Of RK chemicals", "romvijay bio tech (p) ltd.": "ROMVIJAY BIOO TECH PVT LTD", "rukcho biotech": "Rukcho Biotech", "sachin bharud": "Global Polyplast", "sagar biotech pvt ltd": "Sagar Biotech Pvt Ltd", "sagar padgilwar": "Pad Corp", "sanchali rawat": "JAIPUR BIO FERTILIZERS", "sandeep kumar raghuwanshi": "Sandeep kumar Raghuwanshi", "sanjay brothers": "Sanjay Brothers", "sanket jagtap": "Sethu Farmer Producer Company Limited", "sarpan seeds": "Sarpan", "sarthak": "SARTHAK", "satish sajjan": "Satish Sajjan", "shanti devi agriculture store matour": "Shanti Devi Agriculture Store Matour", "shine brand seeds": "SHINE BRAND SEEDS", "shiv kailash green energy private limited": "Shiv kailash Green Energy Private Limited", "shree agro agencies": "Shree Agro Agencies", "shree gopalji impex": "SHREE GOPALJI IMPEX", "shree industries": "Shree Industries", "shree sanwariya trading": "SHREE SANWARIYA TRADING", "shriyap mushroom": "1Shriyap Mushroom", "shubham": "Shubham01", "shubham enterprises": "SHUBHAM ENTERPRISES", "sickle innovations private limited": "sickle", "siddharth parikh": "AGREO SOLUTIONS", "siddhi vinayak enterprises": "Siddhi Vinayak Enterprises", "siva reddy": "Hilfiger Chems", "sk agrotech": "SK AGROTECH", "slavs agrotech": "SLAVS AGROTECH1", "sonkul agro industries pvt. ltd.": "Sonkul Agro Industries Pvt. Ltd.", "sri jyotiba fertilizers": "Sri Jyotiba Fertilizers", "sri sai forestry": "New SRI SAI FORESTRY", "sridhar r": "SRIDHAR R", "srinivasan krishnan": "Srinivasan Krishnan", "startek chemicals limited": "1Star Chemicals", "sudhir goyal": "Sudhir Goyal", "sugan chand sunil kumar": "Sugan Chand Sunil Kumar", "suttind seeds": "Suttind Seeds", "thylakoid biotech pvt. ltd.": "Thylakoid Biotech Pvt. Ltd.", "titan agritech limited": "TITAN AGRITECH LIMITED", "torrent crop science": "TORRENT CROP SCIENCE", "tummuru suresh": "tummuru suresh", "turning point natural care": "Turning point natural care", "tushan sharma": "JAI BALAJI MINERALS", "unison engg. industries": "1Unison Engg. Industries", "urja agricare co.": "Urja Agricare Co.", "urja agriculture company": "1Urja Agriculture Company", "utkarsh agrochem pvt. ltd.": "Utkarsh Agrochem", "vahra technology": "VAHRA TECHNOLOGY", "vaibhav singh chhabra of pep solution": "Vaibhav Singh Chhabra of PEP Solution", "vanproz mp": "1Vanproz Agrovate", "vansthali kisan producer company limited": "Vansthali Kisan Producer Company Limited", "vasudha irrigation": "VASUDHA IRRIGATION", "vedant speciality packaging": "Vedant Speciality Packaging", "venus agro chemicals": "VENUS AGRO CHEMICALS", "vetmantra formulations": "Vetmantra Formulations", "vijay makwana": "Kirtiman Agro Genetics Ltd.", "vikalp": "Vikalp Bio", "vikas aggarwal": "Jayesh Enterprises", "villajio technologies private limited": "VILLAJIO TECHNOLOGIES PRIVATE LIMITED", "vinay krishna c of tech source solutions": "Vinay Krishna C", "vishakha adekar": "Vishakha adekar", "vivek malani": "Vivek Malani", "vr international": "VR International", "v-sar enterprise": "V-Sar Enterprise", "yash pardeshi": "Yash Pardeshi", "z. n. global nation": "Z. N. Global Nation", "zeal biologicals": "Zeal Biologicals", "agribegri trade link pvt. ltd.": "Godawon", "noble crop science": "Godawon", "rain bio tech": "Godawon","rain biotech industries private limited":"RAIN BIOTECH INDUSTRIES PRIVATE LIMITED", "atpl": "Real Trust Exim Corporation, India .","kuldeep singh":"Kuldeep Singh","pawan dhull pesticides":"PAWAN DHULL PESTICIDES","ms poonam auti":"Ms Poonam Auti","sarvin agro chemicals private limited":"SARVIN AGRO CHEMICALS PRIVATE LIMITED","shri paliwal agro care":"SHRI PALIWAL AGRO CARE","balaji trading co.":"Pawan kumar","hifield organics inc.":"Hifield-AG" }

# =====================================================
# GET PICKUP NAME
# =====================================================

def get_pickup_name(order):

    seller_name = (
        order.get("seller_name", "")
        .strip()
        .lower()
    )

    company_name = (
        order.get("company_name", "")
        .strip()
        .lower()
    )

    logging.info(
        f"Seller Name: {seller_name}"
    )

    logging.info(
        f"Company Name: {company_name}"
    )

    # ============================================
    # SPECIAL CASE FOR ATPL (MATCHING m_happy_flow_final copy.py)
    # ============================================

    if seller_name == "atpl" or seller_name == "aptl":
        if "barrix agro science" in company_name:
            pickup_name = "Barrix Agro Science Pvt. Ltd."
            logging.info(f"ATPL Custom Pickup: {pickup_name}")
            return pickup_name
        elif "neptune fairdeal" in company_name:
            pickup_name = "Neptune"
            logging.info(f"ATPL Custom Pickup: {pickup_name}")
            return pickup_name
        elif "real trust exim" in company_name:
            pickup_name = "Real Trust Exim Corporation, India"
            logging.info(f"ATPL Custom Pickup: {pickup_name}")
            return pickup_name
        
        # Fallback to map if no keyword matches
        if company_name in SELLER_PICKUP_MAP:
            pickup_name = SELLER_PICKUP_MAP[company_name]
            logging.info(f"ATPL Pickup (Map): {pickup_name}")
            return pickup_name

        logging.warning(f"ATPL/APTL company not mapped: {company_name}")
        return None


    # ============================================
    # NORMAL SELLER MATCH
    # ============================================

    if seller_name in SELLER_PICKUP_MAP:

        pickup_name = SELLER_PICKUP_MAP[
            seller_name
        ]

        logging.info(
            f"Pickup Matched: {pickup_name}"
        )

        return pickup_name

    logging.warning(
        f"No Pickup Mapping Found "
        f"For Seller: {seller_name}"
    )

    return None

# =====================================================
# PROCESS ORDERS
# =====================================================

def filter_orders(data):

    orders = data.get("response_data", [])

    if not orders:
        logging.info("No Orders Found")
        return []

    logging.info(f"Total Orders Found: {len(orders)}")

    accepted_orders = []
    skipped_orders = []

    for index, order in enumerate(orders, start=1):

        logging.info("===================================")
        logging.info(f"PROCESSING ORDER {index}")
        logging.info("===================================")

        # =================================================
        # GET ORDER DATA
        # =================================================
        order_id = order.get("order_number")
        pickup_name = get_pickup_name(order)
        order["pickup_name"] = pickup_name # Store it in the order object
        weight = float(order.get("weight", 0))
        amount = float(order.get("price", 0))

        logging.info(f"Order ID: {order_id}")
        logging.info(f"Pickup Name: {pickup_name}")
        logging.info(f"Weight: {weight}")
        logging.info(f"Amount: {amount}")

        # =================================================
        # FILTER CONDITIONS
        # =================================================

        reasons = []

        if weight > WEIGHT_LIMIT_GRAM:
            reasons.append("Weight Exceeded")

        if amount > AMOUNT_LIMIT:
            reasons.append("Amount Exceeded")

        # =================================================
        # SKIP ORDER
        # =================================================

        if reasons:

            reason_text = " | ".join(reasons)

            logging.warning(
                f"SKIPPED ORDER {order_id} | {reason_text}"
            )

            skipped_orders.append({
                "order_id": order_id,
                "reason": reasons,
                "weight": weight,
                "amount": amount
            })
            update_history({
                "order_id": order_id,
                "status": "Skipped",
                "seller": order.get("seller_name", "Unknown"),
                "customer": order.get("order_by", "Unknown"),
                "amount": amount,
                "reason": reason_text,
                "bot": "5Bot"
            })

            continue
        # =================================================
        # ACCEPTED ORDER
        # =================================================

        accepted_orders.append(order)

        logging.info(
            f"ORDER ACCEPTED: {order_id}"
        )

    # =================================================
    # SAVE ACCEPTED ORDERS
    # =================================================

    with open(
        "accepted_orders.json",
        "w",
        encoding="utf-8"
    ) as f:

        json.dump(
            accepted_orders,
            f,
            indent=4,
            ensure_ascii=False
        )

    logging.info(
        "accepted_orders.json saved"
    )

    # =================================================
    # SAVE SKIPPED ORDERS
    # =================================================

    with open(
        "skipped_orders.json",
        "w",
        encoding="utf-8"
    ) as f:

        json.dump(
            skipped_orders,
            f,
            indent=4,
            ensure_ascii=False
        )

    logging.info(
        "skipped_orders.json saved"
    )

    # =================================================
    # SUMMARY
    # =================================================

    logging.info("===================================")
    logging.info(f"Accepted Orders: {len(accepted_orders)}")
    logging.info(f"Skipped Orders: {len(skipped_orders)}")
    logging.info("===================================")

    return accepted_orders

# =====================================================
# RUN AUTOMATION
# =====================================================

def run_automation(
    driver,
    wait,
    download_path,
    accepted_orders,
    bot_id
):
    for index, order in enumerate(accepted_orders, start=1):
        try:
            order_id = order.get("order_number")
            
            logging.info("===================================")
            logging.info(f"AUTOMATION FOR ORDER {index}: {order_id}")
            logging.info("===================================")

            # ================================================
            # SEARCH ORDER (LAST TO FIRST SCAN)
            # ================================================

            def go_to_last_page():
                try:
                    # Look for active 'Last' button
                    last_btns = driver.find_elements(By.XPATH, "//a[contains(text(),'Last') and not(contains(@class, 'disabled'))]")
                    if last_btns:
                        driver.execute_script("arguments[0].click();", last_btns[0])
                        logging.info("Moving to Last Page...")
                        time.sleep(4)
                except:
                    pass

            def clear_search():
                try:
                    box = driver.find_element(By.XPATH, "//input[@aria-controls='dyntable']")
                    box.clear()
                    driver.execute_script("arguments[0].value = '';", box)
                    box.send_keys(u'\ue007') # Enter to reset
                    time.sleep(3)
                except:
                    pass

            def perform_search(oid):
                try:
                    box = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//input[@aria-controls='dyntable']"))
                    )
                    box.clear()
                    driver.execute_script("arguments[0].value = '';", box)
                    box.send_keys(oid)
                    box.send_keys(u'\ue007') # Enter
                    time.sleep(4) # More time for table to filter
                    
                    rows_found = driver.find_elements(By.CSS_SELECTOR, "table#dyntable tbody tr")
                    if rows_found:
                        # Check first row text for the order ID
                        if oid in rows_found[0].text:
                            return rows_found
                        if "No matching records found" in rows_found[0].text:
                            return None
                except:
                    pass
                return None

            # 1. CLICK SEARCH BUTTON TO REFRESH TABLE
            try:
                search_btn = driver.find_element(By.ID, "srchSubmit")
                driver.execute_script("arguments[0].click();", search_btn)
                time.sleep(3)
            except:
                pass

            # 2. ALWAYS START FROM LAST PAGE
            go_to_last_page()
            
            rows = None
            while True:
                rows = perform_search(order_id)
                if rows:
                    logging.info(f"Order {order_id} found!")
                    break
                
                # 2. IF NOT FOUND, TRY PREVIOUS PAGE
                logging.info(f"Order {order_id} not found on this page. Checking Previous page...")
                clear_search() # Must clear search to see pagination again
                
                try:
                    # Look for 'Previous' button that is NOT disabled
                    prev_btn = driver.find_element(By.XPATH, "//a[contains(text(), 'Previous') and not(contains(@class, 'disabled'))]")
                    driver.execute_script("arguments[0].click();", prev_btn)
                    time.sleep(4)
                except NoSuchElementException:
                    # Reached first page or no pagination available
                    err = "Order not present (Checked all pages from Last to First)"
                    logging.error(f"Order {order_id} not present (Checked all pages from Last to First).")
                    log_to_excel(order_id, "Failed", err)
                    update_history({
                        "order_id": order_id,
                        "status": "Error",
                        "seller": order.get("seller_name", "Unknown"),
                        "customer": order.get("order_by", "Unknown"),
                        "amount": order.get("price", "0.00"),
                        "error": err,
                        "bot": f"Bot {bot_id}"
                    })
                    break

            if not rows:
                continue # Skip to next order





            # ================================================
            # GET ACTION LINKS
            # ================================================
            # Row 1, Last Column (Actions)
            actions_td = rows[0].find_element(By.XPATH, "./td[last()]")
            
            # Wait for links to appear (sometimes they load slightly after the row)
            links = []
            for _ in range(10):
                links = actions_td.find_elements(By.TAG_NAME, "a")
                if links: break
                time.sleep(0.5)
            
            action_data = {}
            for link in links:
                href = (link.get_attribute("href") or "").strip()
                title = (link.get_attribute("title") or "").strip()
                inner_html = (link.get_attribute("innerHTML") or "").strip()
                
                # Broaden detection logic
                if any(x in title or x in inner_html for x in ["Remark", "comment-img", "fa-book"]):
                    action_data["remark_link"] = href
                elif any(x in title or x in inner_html for x in ["View", "Search", "fa-search"]):
                    action_data["view_link"] = href
                elif any(x in title.lower() or x in inner_html.lower() for x in ["truck", "fa-truck", "shipping", "delivery"]):
                    action_data["truck_link"] = href

            if not action_data:
                logging.warning(f"DEBUG: No action links found for {order_id}. HTML: {actions_td.get_attribute('outerHTML')}")
            else:
                logging.info(f"Action Links Found: {list(action_data.keys())}")

            # ================================================
            # OPEN TRUCK LINK
            # ================================================
            if "truck_link" in action_data:
                logging.info(f"Opening Truck Link for Order {order_id}")
                
                try:
                    # 1. Click to open modal
                    truck_btn = actions_td.find_element(By.CSS_SELECTOR, "a.get_order_id[href*='surface_modal']")
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", truck_btn)
                    time.sleep(1)
                    truck_btn.click()
                except:
                    truck_btn = actions_td.find_element(By.CSS_SELECTOR, "a.get_order_id[href*='surface_modal']")
                    driver.execute_script("arguments[0].click();", truck_btn)
                
                time.sleep(5) # Give it plenty of time to load and populate hidden fields
                
                try:
                    # 2. Select Surface
                    surface_dropdown = wait.until(EC.visibility_of_element_located((By.ID, "serviceSurface")))
                    
                    # Only select if NOT already selected (avoids resetting pickup)
                    current_surface = driver.execute_script("return arguments[0].options[arguments[0].selectedIndex].text.trim();", surface_dropdown)
                    
                    if "Agribegri" not in current_surface:
                        driver.execute_script("""
                            var el = arguments[0];
                            var val = "Agribegri";
                            for(var i=0; i<el.options.length; i++){
                                if(el.options[i].text.trim().toLowerCase().includes(val.toLowerCase())){
                                    el.selectedIndex = i;
                                    break;
                                }
                            }
                            el.dispatchEvent(new Event('change', { bubbles: true }));
                            if(window.jQuery) { window.jQuery(el).trigger('change'); }
                        """, surface_dropdown)
                        logging.info(f"Selected Surface matching 'Agribegri' for {order_id}")
                        time.sleep(5) # Wait for Pickup addresses to load via AJAX
                    else:
                        logging.info(f"'Agribegri Surface' already selected (or match found) for {order_id}")

                    # 3. Select Pickup
                    pickup_name = order.get("pickup_name")
                    if pickup_name:
                        # Wait for options to load (more than just the placeholder)
                        try:
                            WebDriverWait(driver, 10).until(lambda d: len(d.find_element(By.ID, "servicePickupAddress").find_elements(By.TAG_NAME, "option")) > 1)
                        except:
                            logging.warning(f"Pickup options did not load for {order_id} after 10s")
                            
                        pickup_dropdown = driver.find_element(By.ID, "servicePickupAddress")
                        driver.execute_script("""
                            var el = arguments[0];
                            var textToSelect = arguments[1];
                            var found = false;
                            for(var i=0; i<el.options.length; i++){
                                if(el.options[i].text.trim().toLowerCase().includes(textToSelect.toLowerCase())){
                                    el.selectedIndex = i;
                                    found = true;
                                    break;
                                }
                            }
                            el.dispatchEvent(new Event('change', { bubbles: true }));
                            if(window.jQuery) { window.jQuery(el).trigger('change'); }
                            return found;
                        """, pickup_dropdown, pickup_name)
                        logging.info(f"Selected Pickup '{pickup_name}' for {order_id}")
                        time.sleep(2)

                        # ================================================
                        # ATTEMPT SMART SUBMIT

                        # ================================================
                        # We will try to find the form that actually contains the serviceSurface dropdown
                        submit_success = driver.execute_script("""
                            var select = document.getElementById('serviceSurface');
                            if (select && select.form) {
                                var form = select.form;
                                // Ensure order_id is present in THIS form
                                var orderIdField = form.querySelector('[name*="order_id"]') || form.querySelector('[name*="abo_id"]');
                                if (!orderIdField) {
                                    // If not found in form, try to find it globally and copy it
                                    var globalOrderId = document.querySelector('input[name*="order_id"]') || document.querySelector('input[name*="abo_id"]');
                                    if (globalOrderId) {
                                        var hidden = document.createElement('input');
                                        hidden.type = 'hidden';
                                        hidden.name = 'order_id';
                                        hidden.value = globalOrderId.value;
                                        form.appendChild(hidden);
                                    }
                                }
                                
                                // Trigger the submit button inside THIS specific form
                                var btn = form.querySelector('[name="submit_surface"]') || form.querySelector('input[type="submit"]');
                                if (btn) {
                                    btn.click();
                                    return "Form submitted via button click";
                                } else {
                                    form.submit();
                                    return "Form submitted via form.submit()";
                                }
                            }
                            return "ERROR: Surface dropdown or its parent form not found";
                        """)
                        logging.info(f"Submit Action Result for {order_id}: {submit_success}")

                        # Wait for popup and validate message
                        try:
                            # 1. Wait for the message text
                            msg_el = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "popup_message")))
                            msg_text = msg_el.text.strip()
                            
                            # 2. Check if it's successful
                            if "successfully" in msg_text.lower():
                                logging.info(f"SUCCESS: Order {order_id} -> {msg_text}")
                                
                                # Click OK to close success popup
                                popup_ok = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "popup_ok")))
                                driver.execute_script("arguments[0].click();", popup_ok)
                                logging.info(f"Order {order_id} processed successfully")
                                log_to_excel(order_id, "Success")
                                
                                # DEDUCT BALANCE & UPDATE HISTORY
                                with excel_lock:
                                    balance = get_balance()
                                    if balance >= COST_PER_ORDER:
                                        new_balance = balance - COST_PER_ORDER
                                        update_balance(new_balance)
                                        update_history({
                                            "order_id": order_id,
                                            "status": "Processed",
                                            "seller": order.get("seller_name", "Unknown"),
                                            "customer": order.get("order_by", "Unknown"),
                                            "amount": order.get("price", "0.00"),
                                            "bot": f"Bot {bot_id}"
                                        })
                                        logging.info(f"Balance Deducted. New Balance: {new_balance}")
                                    else:
                                        logging.warning(f"Insufficient balance to deduct for order {order_id}")

                                # ================================================
                                # OPEN DELHIVERY & LOGIN (EXACT XPATHS)
                                # ================================================
                                try:
                                    main_tab = driver.current_window_handle
                                    delhivery_tab = None
                                    for handle in driver.window_handles:
                                        driver.switch_to.window(handle)
                                        if "delhivery.com" in driver.current_url:
                                            delhivery_tab = handle
                                            break
                                    
                                    if not delhivery_tab:
                                        logging.info("Opening Delhivery tab...")
                                        driver.execute_script("window.open('https://one.delhivery.com/v2/login','_blank')")
                                        time.sleep(2)
                                        driver.switch_to.window(driver.window_handles[-1])
                                        delhivery_tab = driver.current_window_handle
                                    else:
                                        driver.switch_to.window(delhivery_tab)
                                        if "login" in driver.current_url:
                                            driver.refresh()
                                            time.sleep(2)

                                    # Multi-Step Login (ONLY if on login page)
                                    current_url = driver.current_url
                                    logging.info(f"Delhivery current URL: {current_url}")
                                    
                                    if "/home" in current_url or "delhivery.com" in current_url and "login" not in current_url:
                                        logging.info("Already logged in — skipping login step")
                                    elif "login" in current_url:

                                        wait_dv = WebDriverWait(driver, 20)
                                        
                                        # 1. Email Step
                                        email_input = wait_dv.until(EC.visibility_of_element_located((By.XPATH, "/html/body/div[1]/div/div/div[2]/div[2]/form/div[1]/div/div[1]/section/input")))
                                        email_input.clear()
                                        email_input.send_keys("complain@agribegri.com")
                                        
                                        continue_btn = wait_dv.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/div[2]/div[2]/form/button")))
                                        driver.execute_script("arguments[0].click();", continue_btn)
                                        logging.info("Email entered and Continue clicked")
                                        time.sleep(2)

                                        # 2. Password Step
                                        password_input = wait_dv.until(EC.visibility_of_element_located((By.XPATH, "/html/body/div/div/div[2]/div[2]/div[2]/form/div[2]/div/div/section/input")))
                                        password_input.send_keys("Agribegri@CL#26")
                                        
                                        login_btn = wait_dv.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/div[2]/div[2]/div[2]/form/button")))
                                        driver.execute_script("arguments[0].click();", login_btn)
                                        logging.info("Password entered and Login clicked")
                                        time.sleep(5) # Wait for redirect to start

                                    # Optimized Dashboard Detection
                                    logging.info("Checking Delhivery interface state...")
                                    dashboard_detected = False
                                    
                                    # FAST PATH: If already on a valid page with menu, skip wait
                                    if "delhivery.com" in driver.current_url and "login" not in driver.current_url:
                                        try:
                                            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//*[contains(@class,'ap-menu-trigger')]")))
                                            logging.info("Delhivery interface ready — skipping dashboard load wait")
                                            dashboard_detected = True
                                        except:
                                            pass

                                    if not dashboard_detected:
                                        for i in range(5):
                                            try:
                                                # Allow any delhivery page except login
                                                WebDriverWait(driver, 10).until(lambda d: "delhivery.com" in d.current_url and "login" not in d.current_url)
                                                
                                                # Step 2: Wait for Menu
                                                WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, "//*[contains(@class,'ap-menu-trigger')]")))
                                                logging.info(f"Dashboard detected successfully (Attempt {i+1})")
                                                
                                                dashboard_detected = True
                                                break
                                            except Exception as wait_e:
                                                logging.info(f"Dashboard load check failed (Attempt {i+1})")
                                                if i < 4: 
                                                    logging.info("Refreshing Delhivery page...")
                                                    driver.refresh()
                                                    time.sleep(3)
                                    
                                    if not dashboard_detected:
                                        raise Exception("Could not detect Delhivery interface after multiple attempts")


                                    # Click Domestic/Surface Dropdown (Trigger) - Robust relative selector for 'Direct Step'
                                    logging.info("Attempting direct step to Domestic Dropdown...")
                                    domestic_dropdown = WebDriverWait(driver, 30).until(
                                        EC.element_to_be_clickable((By.XPATH, "//button[.//i[contains(@class,'fa-truck')]]"))
                                    )

                                    # Try multiple click strategies for Vue/React components
                                    try:
                                        domestic_dropdown.click()
                                        logging.info("Domestic dropdown clicked (native)")
                                    except:
                                        try:
                                            from selenium.webdriver.common.action_chains import ActionChains
                                            ActionChains(driver).move_to_element(domestic_dropdown).click().perform()
                                            logging.info("Domestic dropdown clicked (ActionChains)")
                                        except:
                                            driver.execute_script("arguments[0].click();", domestic_dropdown)
                                            logging.info("Domestic dropdown clicked (JS)")

                                    time.sleep(2)

                                    
                                    # Select AGRIBEGRI SURFACE (Using your exact Xpath)
                                    try:
                                        surface_xpath = "/html/body/div[1]/div/div/div[2]/div/div[2]/div[2]/div[2]/div[2]/div/ul/li[3]/button/div[1]/div[1]/div/div[2]/div[2]"
                                        agribegri_surface = WebDriverWait(driver, 20).until(
                                            EC.element_to_be_clickable((By.XPATH, surface_xpath))
                                        )
                                        driver.execute_script("arguments[0].click();", agribegri_surface)
                                        logging.info("AGRIBEGRI SURFACE selected using exact Xpath")
                                    except Exception as e:
                                        logging.warning(f"Exact Xpath failed, trying text-based fallback: {e}")
                                        agribegri_surface = WebDriverWait(driver, 10).until(
                                            EC.element_to_be_clickable((By.XPATH, "//div[text()='AGRIBEGRI SURFACE']"))
                                        )
                                        driver.execute_script("arguments[0].click();", agribegri_surface)
                                        logging.info("AGRIBEGRI SURFACE selected using fallback")
                                    
                                    time.sleep(2)

                                    # ================================================
                                    # SELECT AWB -> OrderID
                                    # ================================================
                                    logging.info("Opening AWB dropdown...")
                                    awb_dropdown = WebDriverWait(driver, 20).until(
                                        EC.element_to_be_clickable((By.XPATH, "//*[@id='app']/div/div/div[2]/div/div[2]/section/div/div/div[1]/div/label/div/button/div"))
                                    )
                                    driver.execute_script("arguments[0].click();", awb_dropdown)
                                    time.sleep(1)

                                    logging.info("Selecting OrderID mode...")
                                    order_id_option = WebDriverWait(driver, 20).until(
                                        EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/div[2]/div/div[2]/section/div/div/div[1]/div/label/div/div/div/ul/li[3]/button/div[1]/div[1]/span[2]"))
                                    )
                                    driver.execute_script("arguments[0].click();", order_id_option)
                                    logging.info("OrderID search mode selected")
                                    time.sleep(2)

                                    
                                    # ================================================
                                    # SEARCH ORDER ID (WITH RETRY)
                                    # ================================================
                                    search_success = False
                                    for attempt in range(3):
                                        try:
                                            logging.info(f"Pasting Order ID {order_id} (Attempt {attempt+1})...")
                                            search_input = WebDriverWait(driver, 15).until(
                                                EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div/div[2]/div/div[2]/section/div/div/div[2]/div/label/div/div[2]/input"))
                                            )
                                            search_input.clear()
                                            search_input.send_keys(order_id)
                                            time.sleep(2) # Give results time to appear
                                            
                                            # Wait for result to populate and click the correct one
                                            logging.info(f"Waiting for search result for {order_id}...")
                                            result_box = WebDriverWait(driver, 15).until(
                                                EC.element_to_be_clickable((
                                                    By.XPATH,
                                                    f"//div[contains(@class, 'cursor-pointer')][.//span[contains(text(), '{order_id}')]]"
                                                ))
                                            )
                                            driver.execute_script("arguments[0].click();", result_box)
                                            logging.info(f"Order {order_id} found and clicked in Delhivery")

                                            
                                            # ================================================
                                            # PARTIAL PAYMENT & PRINT LABEL
                                            # ================================================

                                            time.sleep(3)
                                            payment_status = order.get("payment_status", "").lower()
                                            cod_amount = order.get("cod_payment_amount", "")
                                            
                                            if "partial" in payment_status and cod_amount:
                                                logging.info(f"Partial Payment detected. Updating amount to: {cod_amount}")
                                                
                                                # Click Pen Icon (Edit)
                                                pen_icon = WebDriverWait(driver, 15).until(
                                                    EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/div[3]/div[2]/div/div/div/div[4]/div[1]/div/div[2]/div/article/div[1]/div/div[2]/button"))
                                                )
                                                driver.execute_script("arguments[0].click();", pen_icon)
                                                time.sleep(1)
                                                
                                                # Enter amount
                                                amount_input = WebDriverWait(driver, 10).until(
                                                    EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div/div[3]/div[2]/div/div/div/div[4]/div[1]/div/div[2]/div/article/div[1]/div/div[2]/div[1]/div/div/div/div/label/div/div[2]/input"))
                                                )
                                                amount_input.clear()
                                                amount_input.send_keys(cod_amount)
                                                
                                                # Click Tick Icon (Save)
                                                tick_btn = WebDriverWait(driver, 10).until(
                                                    EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/div[3]/div[2]/div/div/div/div[4]/div[1]/div/div[2]/div/article/div[1]/div/div[2]/div[2]/button[1]"))
                                                )
                                                driver.execute_script("arguments[0].click();", tick_btn)
                                                logging.info("Partial amount updated successfully")
                                                time.sleep(1)
                                            
                                            # ================================================
                                            # PRINT & DOWNLOAD (WITH RETRY & RENAME)
                                            # ================================================
                                            download_success = False
                                            for dl_attempt in range(2):
                                                try:
                                                    logging.info(f"Clicking Print Shipping Label (Attempt {dl_attempt+1})...")
                                                    print_btn = WebDriverWait(driver, 15).until(
                                                        EC.element_to_be_clickable((By.CSS_SELECTOR, "button.ucp__oms__print-label__button"))
                                                    )
                                                    
                                                    # Get initial file list
                                                    before_files = set(os.listdir(download_path))
                                                    
                                                    # Try multiple click strategies for Vue components
                                                    try:
                                                        print_btn.click()
                                                        logging.info("Print button clicked (native)")
                                                    except:
                                                        try:
                                                            from selenium.webdriver.common.action_chains import ActionChains
                                                            ActionChains(driver).move_to_element(print_btn).click().perform()
                                                            logging.info("Print button clicked (ActionChains)")
                                                        except:
                                                            driver.execute_script("arguments[0].click();", print_btn)
                                                            logging.info("Print button clicked (JS)")

                                                    
                                                    # Wait for new file to appear
                                                    new_file = None
                                                    for _ in range(30): # 30 second timeout
                                                        after_files = set(os.listdir(download_path))
                                                        diff = after_files - before_files
                                                        if diff:
                                                            candidate = list(diff)[0]
                                                            if not candidate.endswith(".crdownload"):
                                                                new_file = candidate
                                                                break
                                                        time.sleep(1)
                                                    
                                                    if new_file:
                                                        old_path = os.path.join(download_path, new_file)
                                                        new_path = os.path.join(download_path, f"{order_id}.pdf")
                                                        
                                                        # Handle if file already exists
                                                        if os.path.exists(new_path): os.remove(new_path)
                                                        
                                                        os.rename(old_path, new_path)
                                                        
                                                        # POST-RENAME VERIFICATION
                                                        if os.path.exists(new_path):
                                                            logging.info(f"PDF verified in folder: {order_id}.pdf")
                                                            download_success = True
                                                            break
                                                        else:
                                                            logging.warning(f"File renamed but NOT found on disk: {order_id}.pdf")
                                                    else:
                                                        logging.warning(f"Download verification failed on attempt {dl_attempt+1}")

                                                        
                                                except Exception as dl_e:
                                                    logging.warning(f"Print click failed: {dl_e}")
                                            
                                            if not download_success:
                                                logging.error(f"FATAL: PDF for {order_id} could not be downloaded after 2 attempts")
                                            
                                            # PRE-CHECK: if already downloaded, skip
                                            existing_pdf = os.path.join(download_path, f"{order_id}.pdf")
                                            if os.path.exists(existing_pdf):
                                                logging.info(f"PDF already exists for {order_id} — skipping re-download")
                                                download_success = True


                                            search_success = True
                                            break
                                        except Exception as e:
                                            logging.warning(f"Delhivery order processing attempt {attempt+1} failed: {e}")
                                            if attempt < 2: time.sleep(2)


                                    
                                    if not search_success:
                                        logging.error(f"Order {order_id} NOT FOUND in Delhivery after 3 attempts")

                                    time.sleep(2)
                                    driver.switch_to.window(main_tab)
                                    
                                    # ================================================
                                    # UPDATE ORDER STATUS: Packed + CL Surface
                                    # ================================================
                                    if download_success and "view_link" in action_data:
                                        try:
                                            logging.info(f"Opening View link in NEW TAB to update status for {order_id}...")
                                            driver.execute_script(f"window.open('{action_data['view_link']}', '_blank');")
                                            time.sleep(2)
                                            driver.switch_to.window(driver.window_handles[-1])
                                            
                                            # Set Order Status to "Packed"
                                            status_select = WebDriverWait(driver, 15).until(
                                                EC.presence_of_element_located((By.ID, "abo_status"))
                                            )
                                            Select(status_select).select_by_value("Packed")
                                            logging.info("Order Status set to: Packed")
                                            time.sleep(1)
                                            
                                            # Set Packed Reason to "CL Surface"
                                            reason_select = WebDriverWait(driver, 10).until(
                                                EC.presence_of_element_located((By.ID, "abo_packed_reason"))
                                            )
                                            Select(reason_select).select_by_value("CL Surface")
                                            logging.info("Packed Reason set to: CL Surface")
                                            time.sleep(1)
                                            
                                            # Submit the form (exact selector)
                                            save_btn = WebDriverWait(driver, 10).until(
                                                EC.element_to_be_clickable((By.CSS_SELECTOR, "input[name='update_order_status']"))
                                            )
                                            driver.execute_script("arguments[0].click();", save_btn)
                                            logging.info(f"Submit clicked for {order_id}")
                                            time.sleep(2)
                                            
                                            # Click OK on popup
                                            try:
                                                popup_ok = WebDriverWait(driver, 10).until(
                                                    EC.element_to_be_clickable((By.ID, "popup_ok"))
                                                )
                                                driver.execute_script("arguments[0].click();", popup_ok)
                                                logging.info("Popup OK clicked")
                                                time.sleep(2)
                                            except:
                                                logging.warning("No popup found after submit")
                                            
                                            # Determine if Special Seller
                                            special_sellers = ["agribegri trade link pvt. ltd.", "noble crop science", "rain bio tech", "atpl"]
                                            seller_lower = order.get("seller_name", "").lower()
                                            company_lower = order.get("company_name", "").lower()
                                            is_special = any(s in seller_lower for s in special_sellers)
                                            
                                            # 1. GENERATE INVOICE (ONLY FOR SPECIAL SELLERS)
                                            invoice_path = None
                                            if is_special:
                                                try:
                                                    logging.info(f"Generating invoice for special seller: {seller_lower}")
                                                    before_files_inv = set(os.listdir(download_path))
                                                    invoice_btn = WebDriverWait(driver, 15).until(
                                                        EC.element_to_be_clickable((By.ID, "btnGenerateInvoice"))
                                                    )
                                                    driver.execute_script("arguments[0].click();", invoice_btn)
                                                    
                                                    # Wait for invoice download
                                                    for _ in range(30):
                                                        after_files_inv = set(os.listdir(download_path))
                                                        diff_inv = after_files_inv - before_files_inv
                                                        if diff_inv:
                                                            candidate_inv = list(diff_inv)[0]
                                                            if not candidate_inv.endswith(".crdownload"):
                                                                invoice_path = os.path.join(download_path, candidate_inv)
                                                                logging.info(f"Invoice downloaded: {candidate_inv}")
                                                                break
                                                        time.sleep(1)
                                                except Exception as inv_e:
                                                    logging.warning(f"Invoice generation failed for {order_id}: {inv_e}")

                                            # 2. EMAIL OR UPLOAD
                                            pdf_path = os.path.join(download_path, f"{order_id}.pdf")
                                            
                                            if is_special:
                                                # SPECIAL SELLERS -> SEND EMAIL
                                                recipient_email = None
                                                if any(s in seller_lower for s in ["agribegri trade link pvt. ltd.", "noble crop science", "rain bio tech"]):
                                                    recipient_email = "shipping.agribegri@gmail.com"
                                                elif "atpl" in seller_lower:
                                                    if "barrix" in company_lower: recipient_email = "info@barrix.in"
                                                    elif "neptune" in company_lower: recipient_email = "crop10.order@gmail.com"
                                                    elif "real trust" in company_lower: recipient_email = "realtrustexim24@gmail.com"
                                                
                                                if recipient_email:
                                                    attachments = [pdf_path]
                                                    if invoice_path: attachments.append(invoice_path)
                                                    
                                                    subject = f"Order {order_id} - Label and Invoice"
                                                    body = f"Please find the attached shipping label and invoice for Order ID: {order_id}."
                                                    send_email(recipient_email, subject, body, attachments)
                                                else:
                                                    logging.warning(f"No recipient email found for special seller: {seller_lower} / Company: {company_lower}")
                                            
                                            elif os.path.exists(pdf_path):
                                                # REGULAR SELLERS -> UPLOAD LABEL
                                                try:
                                                    label_input = WebDriverWait(driver, 10).until(
                                                        EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='label_file']"))
                                                    )
                                                    label_input.send_keys(pdf_path)
                                                    logging.info(f"Label uploaded: {order_id}.pdf")
                                                    time.sleep(1)
                                                    
                                                    # Submit label upload form
                                                    upload_btn = WebDriverWait(driver, 10).until(
                                                        EC.element_to_be_clickable((By.CSS_SELECTOR, "input[name='update_order_status']"))
                                                    )
                                                    driver.execute_script("arguments[0].click();", upload_btn)
                                                    logging.info(f"Label upload submitted for {order_id}")
                                                    time.sleep(2)
                                                    
                                                    # Click OK popup again if shown
                                                    try:
                                                        popup_ok2 = WebDriverWait(driver, 5).until(
                                                            EC.element_to_be_clickable((By.ID, "popup_ok"))
                                                        )
                                                        driver.execute_script("arguments[0].click();", popup_ok2)
                                                        time.sleep(1)
                                                    except:
                                                        pass
                                                except Exception as upload_e:
                                                    logging.warning(f"Label upload failed for {order_id}: {upload_e}")
                                            else:
                                                logging.warning(f"PDF not found for upload: {pdf_path}")

                                            # CLOSE VIEW TAB AND RETURN
                                            driver.close()
                                            driver.switch_to.window(main_tab)
                                            time.sleep(1)
                                        except Exception as view_e:
                                            logging.warning(f"View link status update failed for {order_id}: {view_e}")
                                            # Clean up tabs if error occurs
                                            if len(driver.window_handles) > 2: # Agribegri + Delhivery + View
                                                driver.close()
                                            driver.switch_to.window(main_tab)







                                except Exception as dv_error:
                                    logging.error(f"Delhivery Flow Error: {dv_error}")
                                    log_to_excel(order_id, "Failed", f"Delhivery Error: {dv_error}")
                                    update_history({
                                        "order_id": order_id,
                                        "status": "Error",
                                        "seller": order.get("seller_name", "Unknown"),
                                        "customer": order.get("order_by", "Unknown"),
                                        "amount": order.get("price", "0.00"),
                                        "error": f"Delhivery Error: {dv_error}",
                                        "bot": f"Bot {bot_id}"
                                    })
                                    for handle in driver.window_handles:
                                        driver.switch_to.window(handle)
                                        if "agribegri.com" in driver.current_url: break

                            else:
                                logging.error(f"FAILED: Order {order_id} -> Server Error: {msg_text}")
                                log_to_excel(order_id, "Failed", f"Agribegri Server Error: {msg_text}")
                                update_history({
                                    "order_id": order_id,
                                    "status": "Error",
                                    "seller": order.get("seller_name", "Unknown"),
                                    "customer": order.get("order_by", "Unknown"),
                                    "amount": order.get("price", "0.00"),
                                    "error": f"Agribegri Server Error: {msg_text}",
                                    "bot": f"Bot {bot_id}"
                                })
                                # Close the error popup
                                try:
                                    popup_ok = driver.find_element(By.ID, "popup_ok")
                                    driver.execute_script("arguments[0].click();", popup_ok)
                                except:
                                    pass
                            
                        except TimeoutException:
                            logging.warning(f"Success popup not seen for {order_id}")



                    else:
                        logging.warning(f"No pickup name for {order_id}")

                except Exception as inner_e:
                    logging.error(f"Error inside modal for {order_id}: {inner_e}")
                    log_to_excel(order_id, "Failed", f"Modal Error: {inner_e}")
                    update_history({
                        "order_id": order_id,
                        "status": "Error",
                        "seller": order.get("seller_name", "Unknown"),
                        "customer": order.get("order_by", "Unknown"),
                        "amount": order.get("price", "0.00"),
                        "error": f"Modal Error: {inner_e}",
                        "bot": f"Bot {bot_id}"
                    })

            # Go back or cleanup
            if action_data.get("truck_link") and not action_data["truck_link"].startswith("#"):
                driver.back()
                time.sleep(3)
                wait.until(EC.presence_of_element_located((By.ID, "dyntable")))

            time.sleep(1)
            logging.info(f"Order {order_id} Loop Finished")

        except Exception as e:
            logging.error(f"Error processing order {order_id}: {e}")
            log_to_excel(order_id, "Failed", f"Automation Error: {e}")
            update_history({
                "order_id": order_id,
                "status": "Error",
                "seller": order.get("seller_name", "Unknown"),
                "customer": order.get("order_by", "Unknown"),
                "amount": order.get("price", "0.00"),
                "error": f"Automation Error: {e}",
                "bot": f"Bot {bot_id}"
            })
            continue

def worker_bot(bot_id):

    logging.info(f"BOT {bot_id} STARTED")

    while True:

        try:

            driver, wait, download_path = init_driver(bot_id)

            login(driver, wait)

            apply_filter(driver, wait)

            break

        except Exception as e:

            logging.error(
                f"BOT {bot_id} STARTUP FAILED: {e}"
            )

            try:
                driver.quit()
            except:
                pass

            time.sleep(10)

    while True:
        order = order_queue.get()
        try:
            # CHECK BALANCE BEFORE PROCESSING
            with excel_lock:
                if get_balance() < COST_PER_ORDER:
                    logging.error(f"BOT {bot_id} PAUSED: Insufficient balance.")
                    order_queue.task_done()
                    time.sleep(10)
                    order_queue.put(order) # Put it back to try later
                    continue

            order_id = order.get("order_number")
            logging.info(f"BOT {bot_id} PROCESSING {order_id}")
            run_automation(driver, wait, download_path, [order], bot_id)
        except Exception as e:
            logging.error(f"BOT {bot_id} ERROR: {e}")
            try:
                driver.quit()
            except:
                pass
            logging.info(f"RESTARTING BOT {bot_id}")
            time.sleep(5)
            driver, wait, download_path = init_driver(bot_id)
            login(driver, wait)
            apply_filter(driver, wait)
        finally:
            order_queue.task_done()

# =====================================================
# MAIN
# =====================================================

try:
    # START 5 BOTS
    for i in range(5):
        t = threading.Thread(
            target=worker_bot,
            args=(i+1,),
            name=f"Bot_{i+1}",
            daemon=True
        )
        t.start()
    logging.info("5 Bots Started")
    logging.info("LIVE ORDER MONITOR STARTED")

    # 1. HANDLE TARGET ORDER ID IF PROVIDED
    if TARGET_ORDER_ID:
        logging.info(f"Target Order ID provided: {TARGET_ORDER_ID}. Fetching from API...")
        api_data = fetch_orders()
        if api_data:
            all_orders = api_data.get("response_data", [])
            target_order = next((o for o in all_orders if o.get("order_number") == TARGET_ORDER_ID), None)
            if target_order:
                logging.info(f"Target order {TARGET_ORDER_ID} found. Adding to queue.")
                order_queue.put(target_order)
                order_queue.join()
                logging.info(f"Target order {TARGET_ORDER_ID} processing complete.")
            else:
                logging.warning(f"Target order {TARGET_ORDER_ID} not found in API response.")

    # 2. INITIAL FETCH AND PROCESS
    logging.info("Initial check for existing orders...")
    api_data = fetch_orders()
    if api_data:
        orders = api_data.get("response_data", [])
        
        # Filter all existing orders
        accepted_orders = filter_orders({"response_data": orders})
        
        # Populate processed list with ALL currently seen orders to avoid re-processing
        with processed_lock:
            for order in orders:
                processed_orders.add(order.get("order_number"))
        
        if accepted_orders:
            logging.info(f"Processing {len(accepted_orders)} accepted orders from initial fetch...")
            for order in accepted_orders:
                order_queue.put(order)
            
            # Wait for all initial orders to finish
            order_queue.join()
            logging.info("Initial batch processing complete.")
        else:
            logging.info("No accepted orders found in initial fetch.")
    else:
        logging.warning("Initial API fetch failed.")

    # 3. LOOP FOREVER FOR NEW ORDERS
    while True:
        logging.info("Checking API For New Orders...")
        api_data = fetch_orders()

        orders = api_data.get(
            "response_data",
            []
        )

        new_orders = []

        # ============================================
        # FIND ONLY NEW ORDERS
        # ============================================

        for order in orders:

            order_id = order.get(
                "order_number"
            )

            # Skip old processed orders
            with processed_lock:

                if order_id in processed_orders:
                    continue

                processed_orders.add(order_id)

            new_orders.append(order)

            logging.info(
                f"NEW ORDER DETECTED: {order_id}"
            )

        # ============================================
        # PROCESS ONLY NEW ORDERS
        # ============================================

        if new_orders:

            logging.info(
                f"New Orders Found: {len(new_orders)}"
            )

            # ========================================
            # FILTER NEW ORDERS
            # ========================================

            accepted_orders = filter_orders({
                "response_data": new_orders
            })

            # ========================================
            # OVERWRITE JSON FILE
            # ========================================

            with open(
                "accepted_orders.json",
                "w",
                encoding="utf-8"
            ) as f:

                json.dump(
                    accepted_orders,
                    f,
                    indent=4,
                    ensure_ascii=False
                )

            logging.info(
                "accepted_orders.json refreshed"
            )

            # ========================================
            # RUN AUTOMATION
            # ========================================

            if accepted_orders:

                # Search for the first order in the new batch to ensure UI updates
                for order in accepted_orders:

                    order_queue.put(order)

                    logging.info(
                        f"NEW ORDER ADDED TO QUEUE: "
                        f"{order.get('order_number')}"
                    )
                
                # WAIT FOR ALL BOTS TO FINISH THE CURRENT BATCH
                logging.info(f"WAITING: Starting wait for {len(accepted_orders)} orders to finish...")
                order_queue.join()
                logging.info("WAIT FINISHED: All orders in this batch are done. Proceeding to next API check.")

        else:

            logging.info(
                "No New Orders Found"
            )

        # ============================================
        # WAIT BEFORE NEXT API CHECK
        # ============================================

        time.sleep(30)

except Exception as e:

    logging.error(f"MAIN ERROR: {e}")
