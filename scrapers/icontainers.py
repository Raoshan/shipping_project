import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import sys
from datetime import datetime
import pandas as pd
import re
import os
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))            
file_path = os.path.join(BASE_DIR, "global", "rates.xlsx")
print(BASE_DIR)
print(file_path)
columns = [
    "date",
    "ship_line",
    "org_port",
    "dest_port",
    "cont_type",
    "weight",
    "etd_from",
    "eta_to",
    "ocean_frt",
    "frt_surchg",
    "exp_surchg",
    "imp_surchg",
    "total_usd",
    "remarks"
]

origin_port_code = sys.argv[1]
origin_port_name = sys.argv[2]
destination_port_code = sys.argv[3]
destination_port_name = sys.argv[4]
container_size = sys.argv[5]
print("Received arguments:", origin_port_code, origin_port_name, destination_port_code, destination_port_name, container_size)
container_map = {
    "20ft Dry Van": "1 × 20' ST",
    "40ft Dry Van": "1 × 40' ST",
    "40ft Dry Van High Cube": "1 × 40' HC"
}
cont_size=container_map.get(container_size, "")
print("cont size", cont_size)
data = []
# =========================================================
# CHROME SETUP
# =========================================================

options = uc.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument(f"--user-data-dir={os.path.join(BASE_DIR, 'chrome', 'icontainers')}")
options.add_argument(r"--profile-directory=Default")
driver = uc.Chrome(options=options, version_main=148, use_subprocess=True)

wait = WebDriverWait(driver, 30)

# =========================================================
# OPEN WEBSITE
# =========================================================

driver.get("https://my.icontainers.com/sign-in")

driver.maximize_window()

time.sleep(10)

# =========================================================
# CHECK LOGIN
# =========================================================

already_logged_in = False

try:

    dashboard = driver.find_element(
        By.CSS_SELECTOR,
        "div[class='sc-dxjgzI fCktDF']"
    ).text.strip()

    print("DASHBOARD :", dashboard)

    if "Aman" in dashboard:

        already_logged_in = True

        print("Already Logged In ✅")

except:

    print("Login Required 🔐")

# =========================================================
# LOGIN ONLY IF NEEDED
# =========================================================

if not already_logged_in:

    while True:

        try:

            # =============================================
            # USERNAME
            # =============================================

            username = wait.until(
                EC.presence_of_element_located(
                    (
                        By.CSS_SELECTOR,
                        "input[type='email']"
                    )
                )
            )

            username.clear()

            username.send_keys(
                "aman.aggarwal@softwareimpex.com"
            )

            # =============================================
            # PASSWORD
            # =============================================

            password = wait.until(
                EC.presence_of_element_located(
                    (
                        By.CSS_SELECTOR,
                        "input[type='password']"
                    )
                )
            )

            password.clear()

            password.send_keys(
                "4gRmrzs2FPf3Knj"
            )

            # =============================================
            # REMEMBER ME
            # =============================================

            remember = wait.until(
                EC.presence_of_element_located(
                    (
                        By.CSS_SELECTOR,
                        "input[name='remember']"
                    )
                )
            )

            if not remember.is_selected():

                remember.click()

            # =============================================
            # LOGIN BUTTON
            # =============================================

            login_btn = wait.until(
                EC.element_to_be_clickable(
                    (
                        By.CSS_SELECTOR,
                        "button[type='submit']"
                    )
                )
            )

            login_btn.click()

            print("Login Button Clicked ✅")

            time.sleep(20)

            # =============================================
            # CHECK LOGIN SUCCESS
            # =============================================

            dashboard = driver.find_element(
                By.CSS_SELECTOR,
                "div[class='sc-dxjgzI fCktDF']"
            ).text.strip()

            print("DASHBOARD :", dashboard)

            if "Aman" in dashboard:

                print("Login Successful ✅")                
                break

            else:

                print("Login Failed ❌")

        except Exception as e:

            print("Login Error :", e)

            time.sleep(5)

# =========================================================
# LOGIN COMPLETED
# =========================================================

print("READY FOR SCRAPING ✅")
driver.get("https://my.icontainers.com/get-quote/ocean?selectedShipmentType=FCL")
time.sleep(10)
        
origin_input = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[name='origin']")))
origin_input.send_keys(origin_port_code)
origin_options = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div[data-testid='locationDropdownItem']")))
driver.execute_script("arguments[0].click();", origin_options[0])

dest_input = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[name='destination']")))
dest_input.send_keys(destination_port_code)

dest_option = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div[data-testid='locationDropdownItem']")))
time.sleep(2)
driver.execute_script("arguments[0].click();", dest_option[0])
time.sleep(2)


input_box = wait.until(EC.element_to_be_clickable((
    By.XPATH, "//input[contains(@id,'react-select')]"
)))
input_box.click()

# type to trigger dropdown
input_box.send_keys(container_size)

# dynamic xpath
option = wait.until(EC.presence_of_element_located((
    By.XPATH, f"//div[@role='option' and contains(.,'{container_size}')]"
)))

driver.execute_script("arguments[0].click();", option)
commodity_box = wait.until(EC.element_to_be_clickable((
    By.XPATH, "(//input[contains(@id,'react-select')])[2]"
)))
commodity_box.click()
commodity_box.send_keys("Dry or General Cargo")
commodity = wait.until(EC.presence_of_element_located((
    By.XPATH, "//div[@role='option' and contains(.,'Dry or General Cargo')]"
)))
driver.execute_script("arguments[0].click();", commodity)
time.sleep(3)

search_btn = wait.until(EC.element_to_be_clickable((
    By.CSS_SELECTOR, "button[type='submit']"
)))
search_btn.click()
time.sleep(30)

today = datetime.today().strftime('%d-%b-%Y')
cards = []

time.sleep(30)
try:
    all_rates_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-testid='ALL_RATES']")))
    all_rates_btn.click()
except:
    pass
time.sleep(2)
try:
    cards = wait.until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div[data-testid='SingleRate']"))
    )
except:
    data.append([today, "ICONTAINERS", origin_port_name, destination_port_name, cont_size, 20000, "", "", "", "", "", "", "", "No Data"])
    print("Stopping scraper...")
    cards = []

print("Total cards:", len(cards))

try:
    driver.find_element(By.CSS_SELECTOR, "button[aria-label='Close']").click()
    time.sleep(2)
except:
    pass    

import re
from datetime import datetime

import re
from datetime import datetime

for lo in cards:
    try:
        start_date = 'Valid Until'

        raw_end_date = lo.find_element(
            By.XPATH,
            ".//div[.='Valid Until']/following::div[contains(@class,'value')][1]"
        ).text

        end_date = datetime.strptime(raw_end_date, '%d %b %Y').strftime('%d-%b-%Y')
        print(start_date, end_date)

        # open modal
        shows = lo.find_element(By.CSS_SELECTOR, "a[data-testid='ShowRateDetailsButton']")
        driver.execute_script("arguments[0].click();", shows)

        # wait for modal content
        wait.until(EC.visibility_of_element_located(
            (By.XPATH, "//div[.='Origin Charges']")
        ))

        # Origin
        otext = driver.find_element(
            By.XPATH,
            "//div[.='Origin Charges']/ancestor::div[contains(@class,'SectionHeader')]//span"
        ).text
        origin_charge = round(float(re.search(r"[\d,]+\.?\d*", otext).group().replace(",", "")), 2)

        # Freight
        ftext = driver.find_element(
            By.XPATH,
            "//div[.='Freight Charges']/ancestor::div[contains(@class,'SectionHeader')]//span"
        ).text
        freight_charge = round(float(re.search(r"[\d,]+\.?\d*", ftext).group().replace(",", "")), 2)

        export_charge = 0
        import_charge = 0

        total = round(origin_charge + freight_charge + export_charge + import_charge, 2)

        data.append([
            today,
            "ICONTAINERS",
            origin_port_name,
            destination_port_name,
            cont_size,
            20000,
            start_date,
            end_date,
            origin_charge,
            freight_charge,
            export_charge,
            import_charge,
            total,
            ""
        ])

        # close modal
        close_btn = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "button[aria-label='close']")
        ))
        driver.execute_script("arguments[0].click();", close_btn)

        # wait modal close
        wait.until(EC.invisibility_of_element_located(
            (By.XPATH, "//div[.='Origin Charges']")
        ))

    except Exception as e:
        print("Error in one card:", e)
   
new_df = pd.DataFrame(data, columns=columns)
try:
    old_df = pd.read_excel(file_path)
    final_df = pd.concat([old_df, new_df], ignore_index=True)
except:
    final_df = new_df
final_df.to_excel(file_path, index=False)

print("Data saved successfully")

driver.quit()
