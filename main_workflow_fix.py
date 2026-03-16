BALANCE_FILE = "balance.txt"
COST_PER_ORDER = 0.75

def get_balance():
    if not os.path.exists(BALANCE_FILE):
        # Create default balance if not exists
        with open(BALANCE_FILE, "w") as f:
            f.write("0.0")
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

def main_workflow():
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
        print(f" Error reading total count: {e}. Defaulting to single page processing.")
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
                time.sleep(2)
                
                # Call the function for the current index
                result = click_truck_icons_one_by_one(current_index)

                if result == "END":
                    print(f" End of rows on Page {page}.")
                    break

                if result is True:
                    # Deduct balance
                    new_bal = get_balance() - COST_PER_ORDER
                    update_balance(new_bal)
                    print(f" Order processed. Cost {COST_PER_ORDER} deducted. Remaining Balance: {new_bal}")
                    processed_count += 1
                    total_processed_global += 1
                else:
                    print(f" Row {current_index} skipped/incomplete.")
                
                current_index += 1

            except Exception as e:
                print(f" Error processing row {current_index}: {e}")
                current_index += 1 
            
            # ================== CLEANUP TABS ==================
            try:
                while len(driver.window_handles) > 1:
                    driver.switch_to.window(driver.window_handles[-1])
                    driver.close()
                if len(driver.window_handles) > 0:
                    driver.switch_to.window(driver.window_handles[0])
            except:
                pass

        # ================== NEXT PAGE LOGIC ==================
        if page < total_pages:
            print(f" Page {page} done. Moving to Page {page + 1}...")
            try:
                next_btn = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'cus_page_act') and contains(text(), 'Next')]"))
                )
                driver.execute_script("arguments[0].click();", next_btn)
                print(" Waiting for next page to load (stabilizing)...")
                time.sleep(7) 
            except Exception as e:
                print(f" Could not click Next button: {e}")
                break
