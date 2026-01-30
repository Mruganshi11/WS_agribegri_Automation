from selenium import webdriver
import time
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import os
from selenium.webdriver.chrome.service import Service

download_dir = r"C:\Websmith\Agribegri\Downloads"

SELLER_PICKUP_MAP = {
    "rk chemicals gujarat": "dobariya sunil Of RK chemicals",
    "geolife agritech india pvt. ltd maharashtra": "Geolife Agritech India Pvt. Ltd",
    "urja agriculture company": "1Urja Agriculture Company",
    "sickle innovations private limited": "sickle",
    "agribegri trade link pvt. ltd. gujarat": "Godawon",
    "atpl": "Neptune",
    "sagar biotech pvt ltd": "Sagar Biotech Pvt Ltd",
    "essential biosciences": "Essential Biosciences",
    "piyush kataria": "PIYUSH KATARIA"
}


# Set up Chrome options
chrome_options = webdriver.ChromeOptions()
prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": True,

    "safebrowsing.enabled": True
}
chrome_options.add_experimental_option("prefs", prefs)

# Initialize the Chrome driver
# driver = webdriver.Chrome('C:\\Websmith\\Agribegri\\chromedriver.exe', options=chrome_options)
service = Service(r"chromedriver.exe")
driver = webdriver.Chrome(service=service, options=chrome_options)

def login_to_agribegri():
    driver.get('https://agribegri.com/admin/')
    driver.maximize_window()

    username_field = driver.find_element(By.ID, "username").send_keys('Namrata')
    next_button = driver.find_element(By.ID, "btnSubmit")
    next_button.click()
    time.sleep(3)

    password_field = driver.find_element(By.ID, "password").send_keys('Websmith@123456')
    pass_next_button = driver.find_element(By.ID, "btnSubmit")
    pass_next_button.click()
    time.sleep(3)

    enter_otp = driver.find_element(By.ID, "otp").send_keys('123456')
    sign_in = driver.find_element(By.ID, "btnSubmit")
    sign_in.click()
    time.sleep(2)
# def apply_filter():
#     driver.get('https://agribegri.com/admin/manage_orders.php')

#     wait = WebDriverWait(driver, 20)

#     phone_input = wait.until(
#         EC.presence_of_element_located((By.ID, 'srchby_phone'))
#     )
#     phone_input.clear()
#     phone_input.send_keys('6354058079')

#     search_button = driver.find_element(By.ID, 'srchSubmit')
#     search_button.click()

#     # ✅ WAIT until AJAX table reload completes
#     wait.until(
#         EC.presence_of_element_located((By.CSS_SELECTOR, "table tbody tr"))
#     )

#     # ✅ ADD THIS LINE EXACTLY HERE
#     print("Search completed, table is stable")
#     # 🔹 DataTables internal search (top-right Search box)
#     datatable_search = wait.until(
#         EC.element_to_be_clickable(
#             (By.CSS_SELECTOR, "#dyntable_filter input")
#         )
#     )

#     # Clear existing text
#     datatable_search.clear()

#     # Type the value exactly like a human
#     datatable_search.send_keys("RK chemicals GUJARAT")

#     # 🔹 Trigger DataTables filtering (important)
#     driver.execute_script(
#         "arguments[0].dispatchEvent(new Event('keyup'));",
#         datatable_search
#     )

#     # Small wait for table to redraw
#     time.sleep(1)

#     # Ensure filtered rows are present
#     wait.until(
#         EC.presence_of_element_located((By.CSS_SELECTOR, "table tbody tr"))
#     )


#     time.sleep(1)
def apply_filter():
    driver.get('https://agribegri.com/admin/manage_orders.php')

    wait = WebDriverWait(driver, 20)

    phone_input = wait.until(
        EC.presence_of_element_located((By.ID, 'srchby_phone'))
    )
    phone_input.clear()
    phone_input.send_keys('6354058079')

    search_button = driver.find_element(By.ID, 'srchSubmit')
    search_button.click()

    # wait for table load
    wait.until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "table tbody tr"))
    )

    print("Search completed, table is stable")

    time.sleep(1)
   

def click_truck_icons_one_by_one():
    wait = WebDriverWait(driver, 30)

    rows = wait.until(
        EC.presence_of_all_elements_located(
            (By.XPATH, "//table[@id='dyntable']/tbody/tr")
        )
    )

    print(f"Total filtered rows: {len(rows)}")

    #  ONLY FIRST ROW
    row = rows[16]

    # scroll to row
    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center'});", row
    )
    time.sleep(0.5)

    # click truck icon
    truck_icon = row.find_element(By.CSS_SELECTOR, "a.get_order_id")
    driver.execute_script("arguments[0].click();", truck_icon)

    print(" Truck icon clicked for first row")

    # wait for modal
    wait.until(
        EC.visibility_of_element_located(
        (By.CSS_SELECTOR, 'div[data-remodal-id="surface_modal"]')
    )

    )
    print(" Surface modal opened")


    # click dropdown
    surface_dropdown = wait.until(
        EC.element_to_be_clickable((By.ID, "serviceSurface"))
    )
    driver.execute_script("arguments[0].click();", surface_dropdown)
    time.sleep(0.5)

    # click option
    option = driver.find_element(
        By.XPATH,
        "//select[@id='serviceSurface']/option[text()='Agribegri Surface']"
    )
    driver.execute_script("arguments[0].click();", option)

    #  force onchange (CRITICAL)
    driver.execute_script("""
        const select = document.getElementById('serviceSurface');
        select.dispatchEvent(new Event('change', { bubbles: true }));
    """)

    print(" Agribegri Surface selected")


    
    #  correct pickup visibility wait
    wait.until(
        lambda d: d.execute_script(
            "return document.getElementById('li_pickup_address').style.display !== 'none';"
        )
    )


        # ================== PICKUP ADDRESS SELECTION START ==================

    # get seller name from table row (7th column)
    seller_raw = row.find_elements(By.TAG_NAME, "td")[7].text
    seller_name = " ".join(seller_raw.split()).lower()


    print(" Normalized seller name:", seller_name)

    # decide pickup address (mapping + fallback)
    pickup_text = SELLER_PICKUP_MAP.get(seller_name, seller_name)
    print(" Pickup to select:", pickup_text)


    # select pickup address from dropdown
    pickup_dropdown = wait.until(
        EC.element_to_be_clickable((By.ID, "servicePickupAddress"))
    )

    options = pickup_dropdown.find_elements(By.TAG_NAME, "option")

    selected_value = None

    for opt in options:
        if pickup_text.lower() in opt.text.lower():
            selected_value = opt.get_attribute("value")
            print(f" Matching pickup option: {opt.text} (value={selected_value})")
            break

    if not selected_value:
        raise Exception(f" Pickup address not found for seller: {pickup_text}")

    #  SET VALUE DIRECTLY (THIS IS THE FIX)
    driver.execute_script("""
        const select = document.getElementById('servicePickupAddress');
        select.value = arguments[0];
        select.dispatchEvent(new Event('change', { bubbles: true }));
    """, selected_value)

    print("✅ Pickup address selected successfully")


    # ================== PICKUP ADDRESS SELECTION END ==================
    #  RE-SELECT SURFACE AGAIN (CRITICAL FIX)
    driver.execute_script("""
        const surface = document.getElementById('serviceSurface');
        surface.value = 'Agribegri Surface';
        surface.dispatchEvent(new Event('change', { bubbles: true }));
    """)

    print(" Surface re-selected after pickup selection")


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


    # ================== CLICK VIEW ICON ==================

    view_icon = row.find_element(By.CSS_SELECTOR, "img[title='View']")
    driver.execute_script("arguments[0].click();", view_icon)

    print(" View icon clicked (new tab opened)")

    # ================== SWITCH TO VIEW TAB ==================

    driver.switch_to.window(driver.window_handles[-1])
    print(" Switched to order detail tab")


    # ================== COPY ORDER NUMBER (FIXED) ==================
    order_number_elem = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located(
            (By.XPATH, "//span[contains(text(),'Order Number')]/strong")
        )
    )



    order_number = order_number_elem.text.strip()
    print(" Order Number copied:", order_number)


    # ================== OPEN DELHIVERY ==================

    driver.execute_script("window.open('https://one.delhivery.com/home','_blank')")
    driver.switch_to.window(driver.window_handles[-1])
    print("🚚 Delhivery tab opened")

    # ================== DELHIVERY LOGIN ==================


    # 1️ Enter email
    email_input = wait.until(
        EC.visibility_of_element_located((By.NAME, "email"))
    )

    email_input.clear()
    email_input.send_keys("complain@agribegri.com")

    print(" Email entered")

    # 2️ Click Continue
    time.sleep(1.5)

    continue_btn = wait.until(
        EC.presence_of_element_located(
            (By.XPATH, "//button[contains(text(),'Continue')]")
        )
    )

    driver.execute_script("arguments[0].click();", continue_btn)



    print(" Continue clicked")

    # 3️ Wait for password field to appear
    time.sleep(1.5)

    password_input = wait.until(
        EC.visibility_of_element_located((By.XPATH, "//input[@type='password']"))
    )

    password_input.send_keys("AGRIBEGRI!@#26SURFACE")

    print(" Password entered")

    # 5️ Click Login button (Keycloak page)
    login_btn = wait.until(
        EC.presence_of_element_located((By.ID, "kc-login"))
    )

    driver.execute_script("arguments[0].click();", login_btn)

    print("✅ Login button clicked")


    # ================== SELECT AGRIBEGRI SURFACE (TOP-RIGHT DROPDOWN) ==================

    # wait for Delhivery dashboard to load
    WebDriverWait(driver, 30).until(
        EC.presence_of_element_located(
            (By.XPATH, "//div[contains(text(),'Domestic')]")
        )
    )

    # ================== CLICK DOMESTIC DROPDOWN (CORRECT ELEMENT) ==================
# ================== CLICK DOMESTIC / AGRIBEGRI SURFACE DROPDOWN ==================

    domestic_dropdown = WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((
            By.XPATH,
            "//button[contains(@class,'ap-menu-trigger-root')]"
            "[.//i[contains(@class,'fa-truck')]]"
        ))
    )

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", domestic_dropdown)
    time.sleep(0.5)

    driver.execute_script("arguments[0].click();", domestic_dropdown)
    print("✅ Domestic dropdown clicked")



        # ================== SELECT AGRIBEGRI SURFACE ==================

    agribegri_surface = WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((
            By.XPATH,
            "//button[contains(@class,'ap-menu-item')]"
            "[.//div[text()='AGRIBEGRI SURFACE']]"
        ))
    )

    driver.execute_script("arguments[0].click();", agribegri_surface)
    print("✅ AGRIBEGRI SURFACE selected")

    time.sleep(1.5)

    # ================== CLICK AWB DROPDOWN ==================

    awb_dropdown = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((
            By.XPATH,
            "//button[.//span[text()='AWB']]"
        ))
    )

    driver.execute_script("arguments[0].click();", awb_dropdown)
    print("✅ AWB dropdown opened")

    order_id_option = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((
        By.XPATH,
        "//button[.//span[text()='Order ID']]"
    ))
    )

    driver.execute_script("arguments[0].click();", order_id_option)
    print("✅ Order ID selected")

    order_search_input = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((
        By.XPATH,
        "//input[contains(@placeholder,'ORDER ID')]"
    ))
)

    order_search_input.clear()
    order_search_input.send_keys(order_number)

    print(f"✅ Order ID pasted: {order_number}")

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
    print("✅ Order search result clicked")

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
    print("✅ Print Shipping Label button clicked")

    time.sleep(2)

    # ================== WAIT FOR PDF DOWNLOAD & RENAME ==================

    # def wait_and_rename_pdf(download_dir, order_id, timeout=30):
    #     start_time = time.time()
    #     downloaded_file = None

    #     while time.time() - start_time < timeout:
    #         files = os.listdir(download_dir)

    #         # find completed PDF (ignore .crdownload)
    #         pdfs = [
    #             f for f in files
    #             if f.lower().endswith(".pdf") and not f.lower().endswith(".crdownload")
    #         ]

    #         if pdfs:
    #             # assume latest downloaded file
    #             pdfs.sort(
    #                 key=lambda x: os.path.getmtime(os.path.join(download_dir, x)),
    #                 reverse=True
    #             )
    #             downloaded_file = pdfs[0]
    #             break

    #         time.sleep(1)

    #     if not downloaded_file:
    #         raise Exception("❌ PDF download timed out")

    #     old_path = os.path.join(download_dir, downloaded_file)
    #     new_filename = f"{order_id}.pdf"
    #     new_path = os.path.join(download_dir, new_filename)

    #     # if file with same name exists, remove it
    #     if os.path.exists(new_path):
    #         os.remove(new_path)

    #     os.rename(old_path, new_path)
    #     print(f"✅ PDF renamed to: {new_filename}")

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
            raise Exception("❌ PDF download timed out")

        old_path = os.path.join(download_dir, downloaded_file)
        new_path = os.path.join(download_dir, f"{order_id}.pdf")

        if os.path.exists(new_path):
            os.remove(new_path)

        os.rename(old_path, new_path)
        print(f"✅ PDF renamed to: {order_id}.pdf")


        return new_path


    # call the function
    wait_and_rename_pdf(download_dir, order_number)

    # ================== SWITCH BACK TO AGRIBEGRI ADMIN TAB ==================

    for handle in driver.window_handles:
        driver.switch_to.window(handle)
        if "agribegri.com/admin/edit_order.php" in driver.current_url:
            print("✅ Switched back to Agribegri edit order page")
            break


def main_workflow():
    login_to_agribegri()
    apply_filter()

    print("Filtering done. Now clicking truck icon...")

    click_truck_icons_one_by_one()

    print("First row processed: truck clicked + surface selected.")

    input("Press ENTER to close browser...")


main_workflow()
