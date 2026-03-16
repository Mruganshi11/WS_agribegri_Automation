import re
import os
import math
import sys
import json
import time
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException, StaleElementReferenceException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains

sys.path.append(os.path.dirname(os.path.abspath(__file__)))
import email_helper

# PLACEHOLDERS (User to fill)
SENDER_EMAIL = "dispatch.agribegri@gmail.com"
SENDER_PASSWORD = "nxxn wqse vhdn rcwr"
ROW_LIMIT = 0  # 0 for ALL rows, or set a number (e.g. 5)
download_dir = os.path.abspath(r"Downloads")

SELLER_PICKUP_MAP = {
    "rk chemicals gujarat": "dobariya sunil Of RK chemicals",
    "geolife agritech india pvt. ltd maharashtra": "Geolife Agritech India Pvt. Ltd",
    "urja agriculture company delhi": "1Urja Agriculture Company",
    "sickle innovations private limited gujarat": "sickle",
    "agribegri trade link pvt. ltd. gujarat": "Godawon",
    "atpl": "Neptune",
    "sagar biotech pvt ltd": "Sagar Biotech Pvt Ltd",
    "essential biosciences": "Essential Biosciences",
    "piyush kataria rajasthan": "PIYUSH KATARIA",
    "agribegri trade link pvt. ltd.":"Godawon"
}


# Set up Chrome options
current_dir = os.path.dirname(os.path.abspath(__file__))
chrome_driver_path = os.path.join(current_dir, "chromedriver.exe")

chrome_options = webdriver.ChromeOptions()
prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": True,

    "safebrowsing.enabled": True
}
chrome_options.add_experimental_option("prefs", prefs)

service = Service(chrome_driver_path)
driver = webdriver.Chrome(service=service, options=chrome_options)

def login_to_agribegri():
    driver.get('https://agribegri.com/admin/')
    driver.maximize_window()

    username_field = driver.find_element(By.ID, "username").send_keys('Namrata')
    next_button = driver.find_element(By.ID, "btnSubmit")
    next_button.click()
    time.sleep(3)

    password_field = driver.find_element(By.ID, "password").send_keys('Websmith#123456')
    pass_next_button = driver.find_element(By.ID, "btnSubmit")
    pass_next_button.click()
    time.sleep(3)

    enter_otp = driver.find_element(By.ID, "otp").send_keys('123456')
    sign_in = driver.find_element(By.ID, "btnSubmit")
    sign_in.click()
    time.sleep(2)

def apply_filter():
    driver.get('https://agribegri.com/admin/manage_orders.php')

    # Wait for page to load completely
    time.sleep(3)

    wait = WebDriverWait(driver, 30)

    print(" Selecting 'Confirm' status filter...")
    
    # -------- DATE FILTER --------
    today = datetime.now()
    thirty_days_ago = today - timedelta(days=30)
    
    # Using YYYY-MM-DD format as seen in the HTML snippet
    from_date_str = thirty_days_ago.strftime("%Y-%m-%d")
    to_date_str = today.strftime("%Y-%m-%d")
    
    print(f" Applying Date Filter: {from_date_str} to {to_date_str}")
    
    try:
        from_date_input = wait.until(EC.presence_of_element_located((By.ID, "from_date")))
        to_date_input = wait.until(EC.presence_of_element_located((By.ID, "to_date")))
        
        # Using JS to set values if direct send_keys fails due to 'readonly'
        driver.execute_script(f"document.getElementById('from_date').value = '{from_date_str}';")
        driver.execute_script(f"document.getElementById('to_date').value = '{to_date_str}';")
        
        # Select 'Confirm' in the date-wise status dropdown if it exists
        try:
            date_status_dropdown = Select(driver.find_element(By.ID, "search_status_date_order"))
            date_status_dropdown.select_by_value("Confirm")
            print(" Date-wise status set to 'Confirm'")
        except:
            pass
            
        print(" Dates entered successfully")
    except Exception as e:
        print(f" Could not fill dates: {e}")

    # Use JS to select 'Confirm' in the hidden multiselect and trigger the UI update
    driver.execute_script("""
        const select = document.getElementById('search_status');
        // Clear previous selections
        for (let i = 0; i < select.options.length; i++) {
            select.options[i].selected = false;
        }
        // Select 'Confirm'
        for (let i = 0; i < select.options.length; i++) {
            if (select.options[i].value === 'Confirm') {
                select.options[i].selected = true;
                break;
            }
        }
        // Trigger multiselect refresh if the plugin is available
        if (typeof $ !== 'undefined' && $.fn.multiselect) {
            $(select).multiselect('refresh');
        }
        select.dispatchEvent(new Event('change', { bubbles: true }));
    """)

    time.sleep(1)

    search_button = driver.find_element(By.ID, 'srchSubmit')
    search_button.click()

    # wait for table load
    wait.until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "table tbody tr"))
    )

    print("Search completed, table is stable")

    # -------- SORT BY DATE (Oldest First) --------
    try:
        print(" Sorting by Date to ensure Oldest First...")
        # Wait for the Date header to be present
        date_header = wait.until(EC.element_to_be_clickable((By.XPATH, "//table[@id='dyntable']//th[text()='Date']")))
        
        # Check current sort status from the class
        # If 'sorting_asc' is oldest first (standard), we ensure it has that class.
        # If the user specifically said 'descending' to get oldest, we'll aim for that.
        # But looking at your snippet, 'sorting_asc' is the target for ascending sort.
        
        current_class = date_header.get_attribute("class")
        if "sorting_asc" not in current_class:
            print(f" Current sort is {current_class}. Clicking to sort...")
            driver.execute_script("arguments[0].click();", date_header)
            time.sleep(3) # Wait for table refresh
            
            # Re-check if it needs another click (some tables start at 'none' -> 'desc' -> 'asc')
            if "sorting_asc" not in date_header.get_attribute("class"):
                print(" Clicking again to reach Ascending state...")
                driver.execute_script("arguments[0].click();", date_header)
                time.sleep(3)
        
        print(f" Final Sort Class: {date_header.get_attribute('class')}")
    except Exception as e:
        print(f" Failed to sort by date: {e}")

    time.sleep(1)
   

def click_truck_icons_one_by_one(row_index):
    wait = WebDriverWait(driver, 30)
    
    # 1. Initialize detail variables for error logging
    order_id_in_row = "Unknown"
    seller_in_row = "Unknown"
    customer_name = "-"
    net_amount = "-"

    try:
        # Re-fetch rows every time to avoid Stale Elements
        rows = wait.until(
            EC.presence_of_all_elements_located(
                (By.XPATH, "//table[@id='dyntable']/tbody/tr")
            )
        )

        print(f"Total filtered rows in table: {len(rows)}")

        # Check bounds
        if row_index >= len(rows):
            print(f" Index {row_index} out of bounds (End of list)")
            return "END"

        # Select the target row
        row = rows[row_index]
        row_text = row.text.lower()
        
        # Capture basics from table immediately
        try:
            order_id_in_row = row.find_elements(By.TAG_NAME, "td")[4].text.strip()
            seller_in_row = " ".join(row.find_elements(By.TAG_NAME, "td")[7].text.split()).lower()
        except:
            pass

    except (IndexError, StaleElementReferenceException):
        return "END"
    except Exception as e:
        return {"status": "Error", "order_id": "Table Error", "error": f"Failed to read table: {str(e)}"}


    # CHECK REMARK
    if "shipping through transport" in row_text:
        print(f" Row {row_index} skipped (Shipping Through Transport)")
        return {"status": "Skipped", "order_id": order_id_in_row, "seller": seller_in_row, "reason": "Shipping through transport"}
    
    print(f" Row {row_index} selected (Valid for processing)")

    # Start logic wrapper for beautiful errors
    try:
        # ================== PAYMENT STATUS CHECK ==================
        is_partial_payment = False
        if "partial" in row_text and "payment" in row_text:
            is_partial_payment = True
            print(" Partial payment detected for this order")
        else:
            print(" Full / Non-partial payment detected")

        # scroll to row
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", row)
        time.sleep(0.5)

        # click truck icon
        truck_icon = row.find_element(By.CSS_SELECTOR, "a.get_order_id")
        driver.execute_script("arguments[0].click();", truck_icon)
        print(" Truck icon clicked for first row")

        # wait for modal
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'div[data-remodal-id="surface_modal"]')))
        print(" Surface modal opened")

        # click dropdown
        surface_dropdown_elem = wait.until(EC.element_to_be_clickable((By.ID, "serviceSurface")))
        
        # Use Select class for more reliable interaction with <select> tags
        surface_select = Select(surface_dropdown_elem)
        
        print(" Selecting 'Agribegri Surface'...")
        try:
            # Try selecting by visible text first
            surface_select.select_by_visible_text("Agribegri Surface")
        except:
            # Fallback to value if text match fails for some reason
            driver.execute_script("document.getElementById('serviceSurface').value = 'Agribegri Surface';")

        # Force UI update events (CRITICAL for many admin panels)
        driver.execute_script("""
            const select = document.getElementById('serviceSurface');
            select.dispatchEvent(new Event('change', { bubbles: true }));
            select.dispatchEvent(new Event('input', { bubbles: true }));
        """)
        
        # Verify selection
        selected_text = surface_select.first_selected_option.text
        print(f" Surface currently selected: {selected_text}")
        
        if "Agribegri Surface" not in selected_text:
             print(" WARNING: Failed to select via standard method. Retrying with direct JS value set.")
             driver.execute_script("""
                const select = document.getElementById('serviceSurface');
                for (let i = 0; i < select.options.length; i++) {
                    if (select.options[i].text.includes('Agribegri Surface')) {
                        select.selectedIndex = i;
                        break;
                    }
                }
                select.dispatchEvent(new Event('change', { bubbles: true }));
             """)

        # correct pickup visibility wait
        wait.until(lambda d: d.execute_script("return document.getElementById('li_pickup_address').style.display !== 'none';"))

        # ================== PICKUP ADDRESS SELECTION START ==================
        # get seller name from table row (7th column)
        seller_raw = row.find_elements(By.TAG_NAME, "td")[7].text
        seller_name = " ".join(seller_raw.split()).lower()
        print(" Normalized seller name:", seller_name)

        # decide pickup address (mapping + fallback)
        pickup_text = SELLER_PICKUP_MAP.get(seller_name, seller_name)
        print(" Target Pickup Name:", pickup_text)

        # select pickup address from dropdown
        pickup_dropdown = wait.until(EC.element_to_be_clickable((By.ID, "servicePickupAddress")))
        options = pickup_dropdown.find_elements(By.TAG_NAME, "option")

        selected_value = None
        matched_text = ""

        # --- STRATEGY 1: DIRECT MAPPING / EXACT CONTAINMENT ---
        for opt in options:
            if pickup_text.lower() in opt.text.lower():
                selected_value = opt.get_attribute("value")
                matched_text = opt.text
                print(f"  Strategy 1 Match: '{matched_text}' (value={selected_value})")
                break

        # --- STRATEGY 2: REMOVE STATE SUFFIXES (common issue) ---
        if not selected_value:
            print("  Strategy 1 failed. Trying Strategy 2 (Remove State Suffix)...")
            states = ["maharashtra", "gujarat", "madhya pradesh", "karnataka", "delhi", "punjab", "haryana", "rajasthan", "uttar pradesh", "tamil nadu", "telangana", "andhra pradesh", "west bengal"]
            cleaned_name = seller_name
            for state in states:
                if cleaned_name.endswith(state):
                    cleaned_name = cleaned_name.replace(state, "").strip()
                    break
            print(f" Cleaned Name: '{cleaned_name}'")
            if len(cleaned_name) > 3:
                 for opt in options:
                    if cleaned_name in opt.text.lower():
                        selected_value = opt.get_attribute("value")
                        matched_text = opt.text
                        print(f"  Strategy 2 Match: '{matched_text}' (value={selected_value})")
                        break

        # --- STRATEGY 3: FIRST WORD MATCH (Last Resort) ---
        if not selected_value:
            print("  Strategy 2 failed. Trying Strategy 3 (First Word Match)...")
            first_word = seller_name.split()[0]
            if len(first_word) > 3:
                for opt in options:
                    if first_word in opt.text.lower():
                        selected_value = opt.get_attribute("value")
                        matched_text = opt.text
                        print(f"  Strategy 3 Match: '{matched_text}' (value={selected_value})")
                        break

        if not selected_value:
            # debug: print available options
            print(" FAILED. Available Options were:")
            for opt in options:
                 if opt.text.strip():
                    print(f" - {opt.text}")
            raise Exception(f" Pickup address not found for seller: {seller_name} (Target: {pickup_text})")

        # SET VALUE DIRECTLY
        try:
            driver.execute_script("arguments[0].click();", pickup_dropdown)
            time.sleep(0.5)
            select_elem = Select(pickup_dropdown)
            select_elem.select_by_value(selected_value)
            print(f" Selected value {selected_value} using Select class")
            driver.execute_script("""
                const select = arguments[0];
                select.dispatchEvent(new Event('change', { bubbles: true }));
                select.dispatchEvent(new Event('input', { bubbles: true }));
            """, pickup_dropdown)
        except Exception as e:
            print(f" Select class failed, trying JS fallback: {e}")
            driver.execute_script("""
                const select = document.getElementById('servicePickupAddress');
                select.value = arguments[0];
                select.dispatchEvent(new Event('change', { bubbles: true }));
                select.dispatchEvent(new Event('input', { bubbles: true }));
            """, selected_value)

        print(" Pickup address selected successfully")
        time.sleep(1)

        # ================== SUBMIT SURFACE FORM ==================

        submit_btn = wait.until(
            EC.element_to_be_clickable((By.NAME, "submit_surface"))
        )

        driver.execute_script("arguments[0].click();", submit_btn)

        print(" Submit button clicked successfully")

        # ================== SUCCESS POPUP HANDLING ==================

        # wait for success popup
        ok_button = wait.until(
            EC.element_to_be_clickable((By.ID, "popup_ok"))
        )

        driver.execute_script("arguments[0].click();", ok_button)

        print(" Success popup OK clicked")

        # optional: small wait for backend processing
        time.sleep(2)

        # ================== OPEN ORDER DETAIL (ONCE) ==================

        view_icon = row.find_element(By.CSS_SELECTOR, "img[title='View']")
        driver.execute_script("arguments[0].click();", view_icon)
        driver.switch_to.window(driver.window_handles[-1])

        agribegri_order_tab = driver.current_window_handle
        print(" Order detail tab opened & stored")

        # ================== READ DIMENSION (CORRECT PLACE) ==================

        import re

        dimension_text = WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located((
                By.XPATH, "//span[contains(@class,'dimension_lbl')]"
            ))
        ).text.strip()

        print(f" Dimension found: {dimension_text}")

        values = re.findall(r"\d+(?:\.\d+)?", dimension_text)

        if len(values) != 3:
            raise Exception(f" Invalid dimension format: {dimension_text}")

        length, breadth, height = values

        print(f" Parsed: L={length}, B={breadth}, H={height}")


        if not is_partial_payment:
            print(" Performing weight check (non-partial payment)")

            WEIGHT_LIMIT_GRAM = 13000


            wait = WebDriverWait(driver, 20)

            # ---------- READ WEIGHT ----------
            weight_elem = wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "span.weight_lbl"))
            )
            weight_per_unit = float(weight_elem.text.replace(",", "").strip())

            # ---------- READ QUANTITY ----------
            product_row = wait.until(
                EC.presence_of_element_located((
                    By.XPATH,
                    "//table[contains(@class,'table')]/tbody/tr"
                ))
            )
            quantity = float(product_row.find_elements(By.TAG_NAME, "td")[7].text.strip())

            # ---------- CALCULATE TOTAL WEIGHT ----------
            total_weight = weight_per_unit * quantity
            print(f" Total weight: {total_weight} g")

            # ---------- SPLIT CALCULATION (ADD HERE ) ----------
            import math
            split_count = math.ceil(total_weight / WEIGHT_LIMIT_GRAM)
            per_split_weight = total_weight / split_count

            print(f" Total shipments required: {split_count}")
            print(f" Per shipment weight: {per_split_weight:.2f} g")



        else:
            print(" Partial payment order skipping weight check")
            split_count = 1
            per_split_weight = 0


        # ================== SWITCH TO VIEW TAB ==================
        driver.switch_to.window(driver.window_handles[-1])
        print(" Switched to order detail tab")

        agribegri_order_tab = driver.current_window_handle
        print(" Agribegri order tab stored")


        # ================== COPY ORDER NUMBER (FIXED) ==================
        order_number_elem = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located(
            (By.XPATH, "//span[contains(text(),'Order Number')]/strong")
        )
        )



        order_number = order_number_elem.text.strip()
        print(" Order Number copied:", order_number)

        # ================== READ NET AMOUNT FROM AGRIBEGRI ==================

        driver.switch_to.window(agribegri_order_tab)

        net_amount_elem = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, "span_abo_net_amt"))
        )

        net_amount = float(net_amount_elem.text.strip())
        print(f" Net Amount from Agribegri: {net_amount}")

        # calculate per-split amount
        per_split_amount = round(net_amount / split_count, 2)
        print(f" Per shipment amount: {per_split_amount}")

        # switch back to Delhivery tab
        driver.switch_to.window(driver.window_handles[-1])


        agribegri_edit_url = driver.current_url
        print(" Stored Agribegri edit URL:", agribegri_edit_url)
    

        # ================== EXTRACT SHIPPING ADDRESS DETAILS ==================

        print(" Extracting shipping address details...")

        # Name
        customer_name = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "span.name_lbl"))).text.strip()

        # Email
        try:
            customer_email = driver.find_element(By.XPATH, "//td[text()='Email :']/following-sibling::td").text.strip()
        except:
            customer_email = ""

        # Address
        customer_address = driver.find_element(By.CSS_SELECTOR, "span.address_lbl").text.strip()

        # Pincode
        customer_pincode = driver.find_element(By.CSS_SELECTOR, "span.zipcode_lbl").text.strip()

        # Phone
        customer_phone = driver.find_element(By.CSS_SELECTOR, "span.phone_lbl").text.strip()

        # City
        customer_city = driver.find_element(By.CSS_SELECTOR, "span.city_lbl").text.strip()

        # Taluka (optional)
        try:
            customer_taluka = driver.find_element(By.CSS_SELECTOR, "span.taluka_lbl").text.strip()
        except:
            customer_taluka = ""

        # District (optional)
        try:
            customer_district = driver.find_element(By.CSS_SELECTOR, "span.district_lbl").text.strip()
        except:
            customer_district = ""

        # State
        customer_state = driver.find_element(By.CSS_SELECTOR, "span.state_lbl").text.strip()

        print(f" Customer Name: {customer_name}")
        print(f" Customer Email: {customer_email}")
        print(f" Customer Address: {customer_address}")
        print(f" Customer Pincode: {customer_pincode}")
        print(f" Customer Phone: {customer_phone}")
        print(f" Customer City: {customer_city}")
        print(f" Customer Taluka: {customer_taluka}")
        print(f" Customer District: {customer_district}")
        print(f" Customer State: {customer_state}")

        # ================== EXTRACT PRODUCT DESCRIPTION ==================

        print(" Extracting product description...")

        # Product description from the table (column 2 - index 1)
        product_row = wait.until(EC.presence_of_element_located((By.XPATH, "//table[contains(@class,'table')]/tbody/tr")))

        # Product name is in the 2nd column (index 1)
        product_description = product_row.find_elements(By.TAG_NAME, "td")[1].text.strip()

        print(f" Product Description: {product_description}")

        # ================== EXTRACT SELLER EMAIL (New) ==================
        print(" Extracting Seller Email...")
        try:
            seller_email = driver.find_element(By.XPATH, "//h5[text()='Seller Details']/following-sibling::table//td[text()='Email :']/following-sibling::td").text.strip()
            print(f" Seller Email found: {seller_email}")
        except:
            seller_email = ""
            print(" Seller Email not found")

        # ================== OPEN DELHIVERY ==================

        driver.execute_script("window.open('https://one.delhivery.com/v2/login','_blank')")
        driver.switch_to.window(driver.window_handles[-1])
        print(" Delhivery tab opened")

        # ================== DELHIVERY LOGIN (WITH AUTO-SKIP) ==================

        print("Checking if already logged into Delhivery...")
        try:
            # Check if we are already on the dashboard (look for Domestic dropdown or similar)
            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//div[contains(text(),'Domestic')]")))
            print(" Already logged in - Skipping login steps")
        except TimeoutException:
            print(" Not logged in. Proceeding with login...")

            # 1 Enter email
            wait.until(EC.visibility_of_element_located((By.NAME, "email"))).send_keys("complain@agribegri.com")
            print(" Email entered")

            # 2 Click Continue
            time.sleep(1.5)
            driver.execute_script("arguments[0].click();", wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(text(),'Continue')]"))))
            print(" Continue clicked")

            # 3 Wait for password field to appear
            time.sleep(1.5)
            wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@type='password']"))).send_keys("AGRIBEGRI!@#26SURFACE")
            print(" Password entered")

            # 5 Click Login button (Keycloak page)
            driver.execute_script("arguments[0].click();", wait.until(EC.presence_of_element_located((By.ID, "kc-login"))))
            print(" Login button clicked")


        # ================== SELECT AGRIBEGRI SURFACE (TOP-RIGHT DROPDOWN) ==================

        # wait for Delhivery dashboard to load
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "//div[contains(text(),'Domestic')]")))


        # ================== CLICK DOMESTIC / AGRIBEGRI SURFACE DROPDOWN ==================

        domestic_dropdown = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class,'ap-menu-trigger-root')][.//i[contains(@class,'fa-truck')]]")))

        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", domestic_dropdown)
        time.sleep(0.5)

        driver.execute_script("arguments[0].click();", domestic_dropdown)
        print(" Domestic dropdown clicked")



        # ================== SELECT AGRIBEGRI SURFACE ==================

        agribegri_surface = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class,'ap-menu-item')][.//div[text()='AGRIBEGRI SURFACE']]")))

        driver.execute_script("arguments[0].click();", agribegri_surface)
        print(" AGRIBEGRI SURFACE selected")

        time.sleep(1.5)
    
        # ================== WEIGHT > 13 KG SPECIAL FLOW ==================

        if (not is_partial_payment) and (total_weight > 13000):
            print(" Weight > 13 KG flow...")

            # -------- EXPAND SIDEBAR FIRST --------
            try:
                sidebar_toggle = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'ap-sidebar-parent--left__toggler')]//i[contains(@class, 'fa-angles-right')]")))
                driver.execute_script("arguments[0].click();", sidebar_toggle)
            except:
                pass # Sidebar already expanded or toggle not found

            # -------- Click Shipments & Pickups (sidebar) --------
            try:
                shipments_menu = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//button[.//i[contains(@class,'fa-box-circle-check')] and contains(., 'Shipments')]")))
                driver.execute_script("arguments[0].click();", shipments_menu)
                time.sleep(2)
            except:
                print(" Shipments menu not clickable")

            # -------- Find and Click Forward Shipments --------
            forward_shipments = None
            attempts = 0
            while not forward_shipments and attempts < 3:
                attempts += 1
                try:
                    forward_shipments = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((
                            By.XPATH,
                            "//a[contains(@class, 'ap-sidebar__item') and .//div[contains(text(), 'Forward Shipments')]]"
                        ))
                    )
                    break
                except TimeoutException:
                    try:
                        shipments_menu = driver.find_element(By.XPATH, "//button[.//i[contains(@class,'fa-box-circle-check')] and contains(., 'Shipments')]")
                        driver.execute_script("arguments[0].click();", shipments_menu)
                        time.sleep(2)
                    except: pass

            if not forward_shipments:
                raise Exception("Failed to navigate to Forward Shipments")

            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", forward_shipments)
            time.sleep(0.5)
            driver.execute_script("arguments[0].click();", forward_shipments)
            print(" Forward Shipments clicked")
            time.sleep(3)

            # --- CLICK CREATE FORWARD SHIPMENT ---
            selectors = [
                ("//button[contains(@class, 'ap-button') and contains(@class, 'blue') and contains(., 'Create Forward Shipment')]", "Method 1"),
                ("//button[@data-action='create-shipment']", "Method 2"),
                ("//button[contains(text(), 'Create Forward Shipment')]", "Method 3")
            ]
            create_forward_btn = None
            for xpath, desc in selectors:
                try:
                    create_forward_btn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, xpath)))
                    print(f" Found via {desc}")
                    break
                except: pass

            if not create_forward_btn:
                raise Exception("Create Forward Shipment button not found")

            driver.execute_script("arguments[0].click();", create_forward_btn)
            time.sleep(3)

            print(" Filling Delhivery shipment form...")

            wait = WebDriverWait(driver, 20)

            # --- SHIPMENT FORM LOOP ---
            for i in range(1, split_count + 1):
                print(f" Creating shipment part {i}/{split_count}")
                time.sleep(3)

                # -------- Order ID / Reference Number --------
                print(" Looking for Order ID input field...")
                try:
                    order_id_input = WebDriverWait(driver, 30).until(
                        EC.presence_of_element_located((
                            By.XPATH,
                            "//input[@placeholder='Enter Order ID / Reference Number']"
                        ))
                    )
                    print(" Order ID input field found!")
                    
                except Exception as e:
                    print(" All selectors failed for Order ID input")
                    raise Exception("Could not find Order ID input field")

                order_id_input.clear()
                split_order_id = f"{order_number} - {i}"
                order_id_input.send_keys(split_order_id)
                print(f" Order ID filled: {split_order_id}")
                time.sleep(0.5)

            #  STOP HERE FOR NOW (as you asked)
            # Later we will:
            # - Fill item count
            # - Fill weight
            # - Submit
            # - Click "Create Another Shipment"

            # -------- SHIPMENT DESCRIPTION --------
            shipment_desc_input = wait.until(
                EC.presence_of_element_located((
                    By.XPATH,
                    "//input[@placeholder='Enter a description of the item']"
                ))
            )
            shipment_desc_input.clear()
            shipment_desc_input.send_keys(product_description)
            print(f" Shipment Description filled: {product_description}")

            time.sleep(0.5)

            # -------- ITEM COUNT (set to 1) --------
            item_count_input = wait.until(
                EC.presence_of_element_located((
                    By.XPATH,
                    "//input[@type='number' and @placeholder='']"
                ))
            )
            item_count_input.clear()
            item_count_input.send_keys("1")
            print(" Item Count set to: 1")

            time.sleep(0.5)

            # -------- PRODUCT CATEGORY --------
            product_category_input = wait.until(
                EC.presence_of_element_located((
                    By.XPATH,
                    "//input[@placeholder='Select Product Category']"
                ))
            )
            product_category_input.click()
            product_category_input.send_keys("Agriculture")
            time.sleep(1)

            # Select the "Agriculture" option from dropdown
            try:
                agri_option = wait.until(
                    EC.element_to_be_clickable((
                        By.XPATH,
                        "//*[contains(text(), 'Agriculture')]"
                    ))
                )
                driver.execute_script("arguments[0].click();", agri_option)
                print(" Product Category selected: Agriculture")
            except:
                print(" Agriculture option not found in dropdown, typed text should remain")

            time.sleep(0.5)

            #shiping value

            shipment_value_input = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((
                    By.XPATH,
                    "//input[@placeholder='Enter Item Value']"
                ))
            )

            shipment_value_input.clear()
            shipment_value_input.send_keys(str(per_split_amount))
            print(f" Shipment Value entered: {per_split_amount}")

            tax_value_input = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((
                    By.XPATH,
                    "//input[@placeholder='Enter Tax Value']"
                ))
            )

            tax_value_input.clear()
            tax_value_input.send_keys("0")
            print(" Tax Value entered: 0")

            total_value_input = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((
                    By.XPATH,
                    "//input[@placeholder='Enter Total Value']"
                ))
            )

            total_value_input.clear()
            total_value_input.send_keys(str(per_split_amount))
            print(f" Total Value set: {per_split_amount}")




            # -------- PAYMENT MODE (SELECT COD) --------
            print(" Setting Payment Mode to COD...")
            
            # Find the Payment Mode dropdown within Payment Details section
            payment_mode_dropdown = wait.until(
                EC.element_to_be_clickable((
                    By.XPATH,
                    "//h3[text()='Payment Details']/ancestor::div[contains(@class, 'ucp__non-oms__card-common-outline')]//button[contains(@class, 'ap-menu-trigger-root')]"
                ))
            )
            
            # Click to open dropdown
            driver.execute_script("arguments[0].click();", payment_mode_dropdown)
            print(" Payment Mode dropdown opened")
            
            time.sleep(1)
            
            # Select COD option
            cod_option = wait.until(
                EC.element_to_be_clickable((
                    By.XPATH,
                    "//button[contains(@class, 'ap-menu-item')]//span[text()='COD']"
                ))
            )
            driver.execute_script("arguments[0].click();", cod_option)
            print(" Payment Mode set to: COD")
            
            time.sleep(1)


            # ================== OPEN FACILITY DROPDOWN ==================

            wait = WebDriverWait(driver, 30)

            # 1 Locate the Facility dropdown (parent clickable div)
            facility_dropdown = wait.until(
                EC.presence_of_element_located((
                    By.XPATH,
                    "//span[normalize-space()='Select Facility']/ancestor::div[contains(@class,'ap-menu-trigger')]"
                ))
            )

            # 2 Scroll into view
            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", facility_dropdown
            )
            time.sleep(0.5)

            # 3 Force click via JS (works even if overlay exists)
            driver.execute_script("arguments[0].click();", facility_dropdown)

            print(" Facility dropdown opened successfully")

            # ================== TYPE PICKUP ADDRESS ==================

            pickup_input = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((
                    By.XPATH,
                    "//input[@placeholder='Search Pickup Locations']"
                ))
            )

            pickup_input.clear()
            pickup_input.send_keys(pickup_text)

            print(f" Typed pickup address: {pickup_text}")

            time.sleep(1.2)   # allow dropdown results to load

            # ================== SELECT PICKUP FROM LIST ==================

            pickup_option = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((
                    By.XPATH,
                    "//button[contains(@class,'ap-menu-item')]"
                    "[.//span[contains(@class,'main') and "
                    "contains(translate(normalize-space(text()),"
                    "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),"
                    f"'{pickup_text.lower()}')]]"
                ))
            )

            driver.execute_script("arguments[0].click();", pickup_option)

            print(f" Pickup facility selected: {pickup_text}")

            # -------- ADD CUSTOMER DETAILS --------
            add_customer_btn = wait.until(
                EC.element_to_be_clickable((
                    By.XPATH,
                    "//div[contains(@class, 'text-cta-primary') and contains(text(), 'Add Customer Details')]"
                ))
            )
            driver.execute_script("arguments[0].click();", add_customer_btn)
            print(" Add Customer Details clicked")

            time.sleep(3)  # Wait for modal to fully load

            # ================== FILL ADD CUSTOMER MODAL ==================
            print(" Filling Add Customer modal...")

            try:
                # -------- SPLIT NAME INTO FIRST AND LAST --------
                name_parts = customer_name.split(maxsplit=1)
                first_name = name_parts[0] if len(name_parts) > 0 else customer_name
                last_name = name_parts[1] if len(name_parts) > 1 else ""

                # -------- FIRST NAME --------
                first_name_input = wait.until(
                    EC.presence_of_element_located((
                        By.XPATH,
                        "//input[@name='customer_first_name' or @placeholder='Enter first name']"
                    ))
                )
                first_name_input.clear()
                first_name_input.send_keys(first_name)
                print(f" First Name filled: {first_name}")

                time.sleep(0.3)

                # -------- LAST NAME --------
                last_name_input = driver.find_element(
                    By.XPATH,
                    "//input[@name='customer_last_name' or @placeholder='Enter last name']"
                )
                last_name_input.clear()
                if last_name:
                    last_name_input.send_keys(last_name)
                    print(f" Last Name filled: {last_name}")
                else:
                    print(" Last Name left empty (no last name in source)")

                time.sleep(0.3)

                # -------- EMAIL (OPTIONAL) --------
                email_input = driver.find_element(
                    By.XPATH,
                    "//input[@name='email' or @type='email' or @placeholder='Enter email ID']"
                )
                email_input.clear()
                if customer_email:
                    email_input.send_keys(customer_email)
                    print(f" Email filled: {customer_email}")
                else:
                    print(" Email left empty (optional)")

                time.sleep(0.3)

                # -------- PHONE NUMBER --------
                phone_input = driver.find_element(
                    By.XPATH,
                    "//input[@name='phone_number' or @placeholder='Enter mobile number']"
                )
                phone_input.clear()
                phone_input.send_keys(customer_phone)
                print(f" Phone Number filled: {customer_phone}")

                time.sleep(0.3)

                # -------- SHIPPING ADDRESS LINE 1 --------
                address_line1_input = driver.find_element(
                    By.XPATH,
                    "//input[@placeholder='Address Line 1']"
                )
                address_line1_input.clear()
                address_line1_input.send_keys(customer_address)
                print(f" Shipping Address Line 1 filled: {customer_address}")

                time.sleep(0.3)

                # -------- SHIPPING ADDRESS LINE 2 (OPTIONAL) --------
                # Combine City, Taluka, District for Address Line 2

                wait = WebDriverWait(driver, 20)

                # ================== BUILD FORMATTED ADDRESS ==================

                address_parts = []

                if customer_city:
                    address_parts.append(f"City : {customer_city}")

                if customer_taluka:
                    address_parts.append(f"Taluka : {customer_taluka}")

                if customer_district:
                    address_parts.append(f"District : {customer_district}")

                address_line2 = " ".join(address_parts)

                # Example output:
                # City : Ahmedabad Taluka : Ahmedabad District : Ahmedabad

                # ================== LOCATE INPUT ==================

                address_line2_input = wait.until(
                    EC.element_to_be_clickable((
                        By.XPATH, "//input[@placeholder='Address Line 2']"
                    ))
                )

                # ================== FOCUS + CLEAR ==================

                address_line2_input.click()
                driver.execute_script("arguments[0].value = '';", address_line2_input)

                # ================== PASTE VALUE (REACT SAFE) ==================

                if address_line2:
                    driver.execute_script(
                        """
                        arguments[0].value = arguments[1];
                        arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                        arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                        """,
                        address_line2_input,
                        address_line2
                    )

                    print(f" Address Line 2 filled: {address_line2}")
                else:
                    print(" Address Line 2 left empty")

                time.sleep(0.3)


                # -------- PINCODE --------
                pincode_input = driver.find_element(
                    By.XPATH,
                    "//input[@placeholder='Enter pincode']"
                )
                pincode_input.clear()
                pincode_input.send_keys(customer_pincode)
                print(f" Pincode filled: {customer_pincode}")

                time.sleep(2)  # Wait for State and City to auto-populate from pincode

                # -------- VERIFY STATE AND CITY AUTO-FILLED --------
                try:
                    state_input = driver.find_element(
                        By.XPATH,
                        "//input[@placeholder='Enter state']"
                    )
                    state_value = state_input.get_attribute("value")
                    print(f" State auto-filled: {state_value if state_value else 'Not filled'}")
                except:
                    print(" State field not found")

                try:
                    city_input = driver.find_element(
                        By.XPATH,
                        "//input[@placeholder='Enter City']"
                    )
                    city_value = city_input.get_attribute("value")
                    print(f" City auto-filled: {city_value if city_value else 'Not filled'}")
                except:
                    print(" City field not found")

                # -------- CLICK "ADD CUSTOMER" BUTTON --------
                try:
                    # Try multiple selectors for the Add Customer button
                    add_customer_btn = None
                    
                    # Try method 1: Exact text match with blue button
                    try:
                        add_customer_btn = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((
                                By.XPATH,
                                "//button[contains(@class, 'blue') and contains(., 'Add Customer')]"
                            ))
                        )
                        print(" Found 'Add Customer' button (method 1)")
                    except:
                        pass

                    # Try method 2: Just text content
                    if not add_customer_btn:
                        try:
                            add_customer_btn = WebDriverWait(driver, 5).until(
                                EC.element_to_be_clickable((
                                    By.XPATH,
                                    "//button[normalize-space()='Add Customer']"
                                ))
                            )
                            print(" Found 'Add Customer' button (method 2)")
                        except:
                            pass

                    # Try method 3: Any button with "Add Customer" text
                    if not add_customer_btn:
                        try:
                            add_customer_btn = WebDriverWait(driver, 5).until(
                                EC.element_to_be_clickable((
                                    By.XPATH,
                                    "//button[contains(text(), 'Add Customer')]"
                                ))
                            )
                            print(" Found 'Add Customer' button (method 3)")
                        except:
                            pass

                    if add_customer_btn:
                        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", add_customer_btn)
                        time.sleep(0.5)
                        driver.execute_script("arguments[0].click();", add_customer_btn)
                        print(" 'Add Customer' button clicked")
                        time.sleep(3)  # Wait for modal to close and customer to be added
                    else:
                        print(" Could not find 'Add Customer' button - modal may need manual submit")

                except Exception as btn_error:
                    print(f" Error clicking Add Customer button: {btn_error}")

            except Exception as e:
                print(f" Error filling customer modal: {e}")
                import traceback
                traceback.print_exc()

            # ================== SELECT PACKAGE TYPE: CARDBOARD BOX ==================
            
            print(" Selecting Package Type...")
            
            try:
                # Wait a bit for the page to update after customer is added
                time.sleep(2)

                # Find and click the Package Type dropdown
                package_type_dropdown = wait.until(
                    EC.element_to_be_clickable((
                        By.XPATH,
                        "//button[.//span[text()='Select Package Type']]"
                    ))
                )
                
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", package_type_dropdown)
                time.sleep(0.5)
                driver.execute_script("arguments[0].click();", package_type_dropdown)
                print(" Package Type dropdown opened")
                
                time.sleep(1)
                
                # Select "Cardboard Box" option
                cardboard_option = wait.until(
                    EC.element_to_be_clickable((
                        By.XPATH,
                        "//button[contains(@class, 'ap-menu-item')]//span[contains(text(), 'Cardboard Box')]"
                    ))
                )
                
                driver.execute_script("arguments[0].click();", cardboard_option)
                print(" Package Type set to: Cardboard Box")
                
                time.sleep(1)
                

            except Exception as pkg_error:
                print(f" Error selecting Package Type: {pkg_error}")
                print(" Package Type may need to be selected manually")

            print(" Delhivery form partially filled. Ready for manual box details entry.")

            # ================== FILL BOX DIMENSIONS (L, B, H) ==================

            wait = WebDriverWait(driver, 20)

            def fill_number_input(el, value):
                driver.execute_script(
                    """
                    arguments[0].focus();
                    arguments[0].value = '';
                    arguments[0].value = arguments[1];
                    arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                    arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                    """,
                    el,
                    value
                )

            # L
            length_input = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='L']"))
            )
            fill_number_input(length_input, length)

            # B
            breadth_input = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='B']"))
            )
            fill_number_input(breadth_input, breadth)

            # H
            height_input = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='H']"))
            )
            fill_number_input(height_input, height)

            print(f" Box dimensions filled  L={length}, B={breadth}, H={height}")



            # ================== FILL PACKAGED WEIGHT ==================

            wait = WebDriverWait(driver, 20)

            # Round weight and ensure integer (Delhivery prefers whole grams)
            packaged_weight = int(round(per_split_weight))

            weight_input = wait.until(
                EC.element_to_be_clickable((
                    By.XPATH,
                    "//input[@placeholder='Enter packaged weight']"
                ))
            )

            driver.execute_script(
                """
                arguments[0].focus();
                arguments[0].value = '';
                arguments[0].value = arguments[1];
                arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                """,
                weight_input,
                packaged_weight
            )

            print(f" Packaged weight filled: {packaged_weight} g")

            create_forward_btn = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((
                    By.XPATH,
                    "//button[normalize-space()='Create Forward Shipment']"
                ))
            )

            driver.execute_script("arguments[0].click();", create_forward_btn)
            print(" Create Forward Shipment clicked")

            time.sleep(3)

            # ================== CLICK PRINT SHIPPING LABEL ==================

            print_label_btn = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((
                    By.XPATH,
                    "//button[normalize-space()='Print Shipping Label']"
                ))
            )

            driver.execute_script("arguments[0].click();", print_label_btn)
            print(" Print Shipping Label button clicked")

            # ================== RENAME DOWNLOADED SHIPPING LABEL ==================

            def wait_and_rename_pdf(download_dir, new_name, timeout=60):
                print(" Waiting for shipping label download...")

                before_files = set(os.listdir(download_dir))
                end_time = time.time() + timeout
                downloaded_pdf = None

                while time.time() < end_time:
                    current_files = set(os.listdir(download_dir))
                    new_files = current_files - before_files

                    for f in new_files:
                        if f.lower().endswith(".pdf") and not f.lower().endswith(".crdownload"):
                            downloaded_pdf = f
                            break

                    if downloaded_pdf:
                        break

                    time.sleep(0.5)

                if not downloaded_pdf:
                    raise Exception(" Shipping label PDF did not download")

                old_path = os.path.join(download_dir, downloaded_pdf)
                safe_name = new_name.replace("/", "_").replace(" ", "_").strip()
                if not safe_name.lower().endswith(".pdf"):
                    safe_name += ".pdf"
                
                new_path = os.path.join(download_dir, safe_name)

                if os.path.exists(new_path):
                    os.remove(new_path)

                os.rename(old_path, new_path)

                print(f" Shipping label renamed: {safe_name}")
                return new_path

            # Rename using the SPLIT order ID (e.g. 12345 - 1.pdf)
            renamed_pdf_path = wait_and_rename_pdf(download_dir, split_order_id)
            
            # ================== SEND EMAIL ==================
            if seller_email:
                print(f" Sending email to {seller_email}...")
                email_subject = f"Shipping Label for Order {split_order_id}"
                email_body = f"Please find attached the shipping label for order {split_order_id}.\n\nProduct: {product_description}\n\nRegards,\nAutomation Script"
                
                email_helper.send_email_with_attachment(
                    SENDER_EMAIL, 
                    SENDER_PASSWORD, 
                    seller_email, 
                    email_subject, 
                    email_body, 
                    renamed_pdf_path
                )
            else:
                print(" Skipping email (no seller email found)")

            # ================== PREPARE FOR NEXT LOOP (USER REQUESTED NAVIGATION) ==================
            # "navigate to that page and again click on Create Forward shipment"
            if i < split_count:
                print(" Preparing for next split shipment (Explicit Navigation)...")
                
                try:
                    # 1. Click Sidebar Toggler (Expand Sidebar)
                    print(" Clicking Sidebar Toggler...")
                    try:
                        sidebar_toggler = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((
                                By.XPATH,
                                "//div[contains(@class, 'ap-sidebar-parent--left__toggler')]"
                            ))
                        )
                        driver.execute_script("arguments[0].click();", sidebar_toggler)
                        print(" Sidebar Toggler clicked")
                        time.sleep(1)
                    except:
                        print(" Sidebar might already be expanded or toggler not found")

                    # 2. Click "Shipments & Pickups" (Dropdown)
                    print(" Clicking 'Shipments & Pickups' menu...")
                    shipments_menu = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((
                            By.XPATH,
                            "//button[contains(@class, 'ap-sidebar__item') and .//span[contains(text(), 'Shipments & Pickups')]]"
                        ))
                    )
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", shipments_menu)
                    time.sleep(0.5)
                    driver.execute_script("arguments[0].click();", shipments_menu)
                    print(" 'Shipments & Pickups' clicked")
                    
                    time.sleep(2) # Wait for dropdown to expand

                    # 3. Click "Forward Shipments" Link
                    print(" Clicking 'Forward Shipments' link...")
                    forward_shipments_link = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((
                            By.XPATH,
                            "//a[contains(@class, 'ap-sidebar__item') and .//div[contains(text(), 'Forward Shipments')]]"
                        ))
                    )
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", forward_shipments_link)
                    time.sleep(0.5)
                    driver.execute_script("arguments[0].click();", forward_shipments_link)
                    print(" 'Forward Shipments' clicked")

                    time.sleep(5) # Wait for page reload

                    # 4. Click Create Forward Shipment Button again
                    print(" Looking for 'Create Forward Shipment' button for next loop...")
                    create_forward_btn_loop = WebDriverWait(driver, 10).until(
                         EC.element_to_be_clickable((
                             By.XPATH, 
                             "//button[contains(@class, 'ap-button') and contains(@class, 'blue') and contains(., 'Create Forward Shipment')]"
                        ))
                    )
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", create_forward_btn_loop)
                    time.sleep(0.5)
                    driver.execute_script("arguments[0].click();", create_forward_btn_loop)
                    print(" Create Forward Shipment button clicked (for next part)")
                    
                    time.sleep(3) # Wait for form to open

                except Exception as e:
                    print(f" Failed to reset form via sidebar navigation: {e}")
                    raise e



        #  STOP HERE IF WEIGHT > 13 KG (FULL PAYMENT ONLY)
        if (not is_partial_payment) and (total_weight > 13000):
            print(" Weight > 13 KG detected")
            print(" Stopping after AGRIBEGRI SURFACE selection")
            print(" Do NOT proceed to Shipments & Pickups / AWB")

            return

        # ================== CLICK AWB DROPDOWN ==================

        awb_dropdown = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((
            By.XPATH,
            "//button[.//span[text()='AWB']]"
        ))
        )

        driver.execute_script("arguments[0].click();", awb_dropdown)
        print(" AWB dropdown opened")

        order_id_option = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((
                By.XPATH, "//button[.//span[text()='Order ID']]"
            ))
        )
        driver.execute_script("arguments[0].click();", order_id_option)
        print(" Order ID selected")

        order_search_input = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((
                By.XPATH, "//input[contains(@placeholder,'ORDER ID') or contains(@placeholder,'AWB')]"
            ))
        )
        order_search_input.clear()
        order_search_input.send_keys(order_number)
        print(f" Order ID pasted: {order_number}")

        # ================== CLICK ORDER SEARCH RESULT ==================
        search_result = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//div[contains(@class,'ucp__global-search__results')]"
                "//div[contains(@class,'cursor-pointer')]"
                "[.//span[text()='" + order_number + "']]"
            ))
        )
        driver.execute_script("arguments[0].click();", search_result)
        print(" Order search result clicked")
        time.sleep(2)

        # ================== PARTIAL PAYMENT HANDLING IN DELHIVERY ==================
        if is_partial_payment:
            print(" Partial payment flow activated in Delhivery")

            # ---------- SWITCH TO AGRIBEGRI TAB ----------
            driver.switch_to.window(agribegri_order_tab)
            driver.refresh()
            time.sleep(3)

            print(" Forced return to Agribegri edit order page")


            # ---------- COPY COD PAYMENT AMOUNT ----------
            cod_amount_elem = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((
                    By.XPATH, "//p[contains(@class,'p_disp_advance_amt') and contains(.,'COD Payment Amount')]"
                ))
            )

            import re
            cod_text = cod_amount_elem.text
            cod_amount = re.search(r"([\d]+\.\d+)", cod_text).group(1)

            print(" Correct COD Amount copied:", cod_amount)

            # ---------- SWITCH BACK TO DELHIVERY ----------
            driver.switch_to.window(driver.window_handles[-1])

            # ---------- CLICK  EDIT ICON (CORRECT ELEMENT) ----------
            edit_payment_btn = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@data-action='edit-payment-mode']"))
            )

            driver.execute_script("arguments[0].click();", edit_payment_btn)
            print(" Payment Details edit icon clicked (correct one)")

            time.sleep(0.8)

            # ---------- FORCE-OPEN PAYMENT MODE DROPDOWN ----------
            payment_dropdown = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(@class,'ap-menu-trigger-root')]"))
            )

            driver.execute_script("arguments[0].click();", payment_dropdown)
            time.sleep(0.5)

            cod_option = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'ucp__order-creation__select-payment__dropdown-item--cod')]"))
            )

            driver.execute_script("arguments[0].click();", cod_option)
            print(" Cash On Delivery selected")
            time.sleep(0.5)


            # ---------- ENTER COLLECTABLE AMOUNT ----------
            collectable_input = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input.input[type='number']"))
            )
            collectable_input.clear()
            collectable_input.send_keys(cod_amount)
            print(f" Collectable amount entered: {cod_amount}")

            # ---------- CLICK  SAVE ----------
            save_btn = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@title='Save']"))
            )
            driver.execute_script("arguments[0].click();", save_btn)
            print(" COD payment saved")

            time.sleep(2)


        # ================== CLICK PRINT SHIPPING LABEL ==================

        print_label_btn = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//button[normalize-space()='Print Shipping Label']"
            ))
        )

        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", print_label_btn)
        time.sleep(0.5)

        driver.execute_script("arguments[0].click();", print_label_btn)
        print(" Print Shipping Label button clicked")

        time.sleep(2)


        def wait_and_rename_pdf(download_dir, order_id, timeout=30):
            start_time = time.time()
            downloaded_file = None

            while time.time() - start_time < timeout:
                files = os.listdir(download_dir)

                pdfs = [
                    f for f in files
                    if f.lower().endswith(".pdf") and not f.lower().endswith(".crdownload")
                ]

                if pdfs:
                    pdfs.sort(
                        key=lambda x: os.path.getmtime(os.path.join(download_dir, x)),
                        reverse=True
                    )
                    downloaded_file = pdfs[0]
                    break

                time.sleep(1)

            if not downloaded_file:
                raise Exception(" PDF download timed out")

            old_path = os.path.join(download_dir, downloaded_file)
            new_path = os.path.join(download_dir, f"{order_id}.pdf")

            if os.path.exists(new_path):
                os.remove(new_path)

            os.rename(old_path, new_path)
            print(f" PDF renamed to: {order_id}.pdf")


            return new_path


        # call the function
        wait_and_rename_pdf(download_dir, order_number)

        # ================== SWITCH BACK TO AGRIBEGRI ADMIN TAB ==================
        driver.switch_to.window(agribegri_order_tab)
        driver.refresh()
        time.sleep(2)

        print(" Switched back to Agribegri edit order page (final)")

        # ================== SET ORDER STATUS TO PACKED ==================

        # wait for Order Status dropdown
        order_status_select = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "abo_status"))
        )

        # use Select for <select> tag
        select = Select(order_status_select)

        # select "Packed"
        select.select_by_visible_text("Packed")

        print(" Order status set to Packed")

        time.sleep(1)

        # ================== SET PACKED REASON TO CL SURFACE ==================

        packed_reason_select = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "abo_packed_reason"))
        )

        packed_reason = Select(packed_reason_select)

        packed_reason.select_by_visible_text("CL Surface")

        print(" Packed reason set to CL Surface")

        time.sleep(1)


        # ================== SET PAYMENT STATUS IF PARTIAL PAYMENT ==================

        if is_partial_payment:
            print(" Setting payment status: Bank Partial Payment Recieved")

            payment_status_select = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.ID, "abo_payment_type"))
            )

            payment_select = Select(payment_status_select)

            payment_select.select_by_visible_text("Bank Partial Payment Recieved")

            print(" Payment status set to Bank Partial Payment Recieved")

            time.sleep(1)
        else:
            print(" Skipping payment status change (not partial payment)")

        time.sleep(1)
        # ================== CLICK SUBMIT BUTTON ==================

        submit_btn = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((
                By.NAME, "update_order_status"
            ))
        )

        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", submit_btn)
        time.sleep(0.5)

        driver.execute_script("arguments[0].click();", submit_btn)
        print(" Submit button clicked successfully")

        time.sleep(2)
        # ================== CLICK OK ON SUCCESS POPUP ==================

        ok_btn = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, "popup_ok"))
        )

        driver.execute_script("arguments[0].click();", ok_btn)
        print(" Success popup OK clicked")

        time.sleep(1.5)

        # ================== UPLOAD LABEL PDF ==================

        label_input = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.NAME, "label_file"))
        )

        label_pdf_path = os.path.join(download_dir, f"{order_number}.pdf")

        if not os.path.exists(label_pdf_path):
            raise Exception(f" Label PDF not found: {label_pdf_path}")

        label_input.send_keys(label_pdf_path)
        print(f" Label PDF uploaded: {label_pdf_path}")

        time.sleep(1)
        # ================== CLICK SUBMIT BUTTON ==================

        submit_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.NAME, "update_manifest_file"))
        )

        submit_btn.click()
        print(" Submit button clicked")


        # ================== SUBMIT LABEL UPLOAD ==================

        time.sleep(1)

        # ================== CLICK OK ON POPUP ==================

        ok_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "popup_ok"))
        )

        ok_btn.click()
        print(" OK button clicked on popup")

        time.sleep(1)
        return {
            "status": "Processed",
            "order_id": order_number,
            "seller": seller_name,
            "customer": customer_name,
            "amount": net_amount,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }

    except TimeoutException as te:
        error_msg = f"Timeout: Could not find element on page (system busy or slow)"
        print(f" ERROR: {error_msg}")
        return {
            "status": "Error",
            "order_id": order_id_in_row,
            "seller": seller_in_row,
            "customer": customer_name if 'customer_name' in locals() else "-",
            "amount": net_amount if 'net_amount' in locals() else "-",
            "error": error_msg,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
    except Exception as e:
        raw_err = str(e)
        # Beautify common errors
        if "surface_modal" in raw_err:
            error_msg = "Order detail modal failed to open"
        elif "serviceSurface" in raw_err:
            error_msg = "Surface selection dropdown not found"
        elif "Order ID input field" in raw_err:
            error_msg = "Order not found or form failed in Delhivery"
        elif "Create Forward Shipment button" in raw_err:
            error_msg = "Delhivery Create Shipment button not found"
        else:
            error_msg = f"System Error: {raw_err.split('Stacktrace:')[0].strip()}"
            
        print(f" ERROR: {error_msg}")
        return {
            "status": "Error",
            "order_id": order_id_in_row,
            "seller": seller_in_row,
            "customer": customer_name if 'customer_name' in locals() else "-",
            "amount": net_amount if 'net_amount' in locals() else "-",
            "error": error_msg,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }

BALANCE_FILE = "balance.txt"
ORDERS_COUNT_FILE = "orders_count.txt"
HISTORY_FILE = "processed_orders.json"
COST_PER_ORDER = 0.75

# Unique ID for the current script execution
RUN_ID = datetime.now().strftime("%Y%m%d_%H%M%S")
RUN_TIME = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def clear_history():
    """Removes all processed orders from the history file and resets counter."""
    try:
        if os.path.exists(HISTORY_FILE):
            os.remove(HISTORY_FILE)
            print(f" SUCCESS: History file '{HISTORY_FILE}' deleted.")
        
        if os.path.exists(ORDERS_COUNT_FILE):
            with open(ORDERS_COUNT_FILE, "w") as f:
                f.write("0")
            print(f" SUCCESS: Orders count reset.")
            
    except Exception as e:
        print(f" ERROR: Failed to clear history/counter: {e}")

def get_balance():
    if not os.path.exists(BALANCE_FILE):
        return 0.0
    try:
        with open(BALANCE_FILE, "r") as f:
            return float(f.read().strip())
    except:
        return 0.0

def update_balance(new_balance):
    try:
        with open(BALANCE_FILE, "w") as f:
            f.write(str(round(new_balance, 2)))
    except Exception as e:
        print(f" Error updating balance file: {e}")

def get_orders_count():
    if not os.path.exists(ORDERS_COUNT_FILE):
        return 0
    try:
        with open(ORDERS_COUNT_FILE, "r") as f:
            return int(f.read().strip())
    except:
        return 0

def update_orders_count(new_count):
    try:
        with open(ORDERS_COUNT_FILE, "w") as f:
            f.write(str(new_count))
    except Exception as e:
        print(f" Error updating orders count file: {e}")

def update_history(entry):
    try:
        # Add run metadata to each entry
        entry["run_id"] = RUN_ID
        entry["run_start_time"] = RUN_TIME
        
        history = []
        if os.path.exists(HISTORY_FILE):
            with open(HISTORY_FILE, "r") as f:
                try:
                    history = json.load(f)
                except:
                    history = []
        
        history.append(entry)
        
        # Keep only latest 1000 records to avoid file bloat
        if len(history) > 1000:
            history = history[-1000:]
            
        with open(HISTORY_FILE, "w") as f:
            json.dump(history, f, indent=4)
    except Exception as e:
        print(f" Error updating history file: {e}")

def is_driver_alive():
    """Helper to check if the selenium driver is still responding."""
    try:
        # Simple health check call
        driver.title
        return True
    except:
        return False

def main_workflow():
    # CHECK FOR COMMAND LINE ARGUMENTS
    if len(sys.argv) > 1:
        if sys.argv[1].lower() == "--clear":
            clear_history()
            return # Stop after clearing

    login_to_agribegri()
    apply_filter()

    print("Filtering done. Checking total orders...")

    # ================== GET TOTAL COUNT ==================
    try:
        total_count_elem = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//*[@id='seller_frm']/div/span[1]"))
        )
        total_text = total_count_elem.text.strip()
        import re
        match = re.search(r"(\d+)", total_text)
        if match:
            total_orders = int(match.group(1))
        else:
            total_orders = 50 
            print(" Could not parse total count, defaulting to 50")
            
        print(f" Total Orders Found: {total_orders}")

    except Exception as e:
        print(f" Error reading total count: {e}. Defaulting to processing.")
        total_orders = 50

    import math
    ORDERS_PER_PAGE = 50
    total_pages = math.ceil(total_orders / ORDERS_PER_PAGE)
    print(f" Total Pages to process: {total_pages}")

    processed_count = 0
    total_processed_global = 0

    # ================== PAGE LOOP ==================
    for page in range(1, total_pages + 1):
        print(f"\n --- STARTING PAGE {page} / {total_pages} ---")
        current_index = 0

        while True:
            # Check Balance before starting an order
            current_balance = get_balance()
            if current_balance < COST_PER_ORDER:
                print(f"!!! INSUFFICIENT BALANCE ({current_balance} INR). Need {COST_PER_ORDER} INR per order.")
                print(" Stopping script. Please recharge via Dashboard.")
                return

            # Check Global Limit
            if ROW_LIMIT > 0 and total_processed_global >= ROW_LIMIT:
                print(f" Global Limit of {ROW_LIMIT} rows reached.")
                return

            print(f"\n --- Processing Row Index: {current_index} (Page {page}) ---")
            
            try:
                # Proper delay before each order
                time.sleep(3)
                
                # Health check: If the browser was closed, stop processing
                if not is_driver_alive():
                    print(" CRITICAL: Browser was closed or lost connection. Stopping automation.")
                    break

                # Call the function for the current index
                result = click_truck_icons_one_by_one(current_index)

                if result == "END":
                    print(f" End of rows on Page {page}.")
                    break

                if isinstance(result, dict):
                    # Always update timestamp for continuity
                    if "timestamp" not in result:
                        result["timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                    if result.get("status") == "Processed":
                        # Deduct balance
                        new_bal = get_balance() - COST_PER_ORDER
                        update_balance(new_bal)
                        
                        # Increment order count
                        new_count = get_orders_count() + 1
                        update_orders_count(new_count)
                        
                        # Log success history
                        update_history(result)
                        
                        print(f" SUCCESS: Order {result['order_id']} processed. Cost {COST_PER_ORDER} deducted. Remaining Balance: {new_bal}")
                        processed_count += 1
                        total_processed_global += 1
                        time.sleep(3) # Post-processing delay
                    
                    elif result.get("status") in ["Skipped", "Error"]:
                        update_history(result)
                        if result.get("status") == "Skipped":
                            print(f" Row {current_index} skipped: {result.get('reason')}")
                        else:
                            print(f" Row {current_index} error: {result.get('error')}")
                    else:
                        print(f" Row {current_index} unknown result status: {result.get('status')}")
                else:
                    print(f" Row {current_index} incomplete or non-dict result.")
                
                current_index += 1

            except Exception as e:
                print(f" Unexpected error in main loop for row {current_index}: {e}")
                current_index += 1 
            
            # ================== CLEANUP TABS ==================
            if is_driver_alive():
                try:
                    while len(driver.window_handles) > 1:
                        driver.switch_to.window(driver.window_handles[-1])
                        driver.close()
                    if len(driver.window_handles) > 0:
                        driver.switch_to.window(driver.window_handles[0])
                except:
                    pass
            else:
                print(" Skipping tab cleanup: Driver is not alive.")
                break

        # ================== NEXT PAGE LOGIC ==================
        if page < total_pages:
            print(f" Page {page} done. Moving to Page {page + 1}...")
            try:
                next_btn = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'cus_page_act') and contains(text(), 'Next')]"))
                )
                driver.execute_script("arguments[0].click();", next_btn)
                print(" Waiting for next page to load (stabilizing)...")
                time.sleep(10) 
            except Exception as e:
                print(f" Could not click Next button: {e}")
                break
        else:
             print(" All pages processed.")
    
    print(f"\n Workflow completed. Total rows processed: {total_processed_global}")


if __name__ == "__main__":
    try:
        main_workflow()
    except Exception as e:
        print(f"\n CRITICAL ERROR: {e}")
        import traceback
        traceback.print_exc()
    finally:
        print("\n Script finished. Closing browser in 5 seconds...")
        time.sleep(5)
        # driver.quit()

