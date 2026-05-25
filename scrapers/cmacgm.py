import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import sys
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime
import pandas as pd
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
data = []
today = datetime.today().strftime('%d-%b-%Y')
container_map = {
    "20' Dry Standard": "1 × 20' ST",
    "40' Dry Standard": "1 × 40' ST",
    "40' Dry High Cube": "1 × 40' HC",
    "45' Dry High Cube": "1 × 45' HC"
}

origin_port_code = sys.argv[1]
origin_port_name = sys.argv[2]
destination_port_code = sys.argv[3]
destination_port_name = sys.argv[4]
container_size = sys.argv[5]
print("Received arguments:", origin_port_code, origin_port_name, destination_port_code, destination_port_name, container_size)
print("Final Container Value:", container_map.get(container_size, ""))
cont_size=container_map.get(container_size, "")

# ========= CHROME SETUP =========
options = uc.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument("--disable-blink-features=AutomationControlled")
driver = uc.Chrome(options=options, version_main=148, use_subprocess=True)
wait = WebDriverWait(driver, 20)

# ================= OPEN WEBSITE =================
driver.get("https://www.cma-cgm.com/")
driver.maximize_window()

time.sleep(5)

# ================= OPEN MY CMA =================
my_cma = wait.until(
    EC.element_to_be_clickable(
        (
            By.XPATH,
            "//a[contains(@class,'mycma')]"
        )
    )
)

driver.execute_script(
    "arguments[0].click();",
    my_cma
)

time.sleep(3)

# ================= LOGIN BUTTON =================
login_btn = wait.until(
    EC.element_to_be_clickable(
        (
            By.XPATH,
            "//button[contains(text(),'Log In')]"
        )
    )
)

driver.execute_script(
    "arguments[0].click();",
    login_btn
)

time.sleep(3)

# ================= EMAIL =================
email = wait.until(
    EC.presence_of_element_located(
        (
            By.ID,
            "login-email"
        )
    )
)

email.clear()

email.send_keys(
    "Ocean@tanuimpex.com"
)

# ================= PASSWORD =================
password = wait.until(
    EC.presence_of_element_located(
        (
            By.ID,
            "login-password"
        )
    )
)

password.clear()

password.send_keys(
    "Mtanu2025#"
)

# ================= FINAL LOGIN =================
final_login = wait.until(
    EC.element_to_be_clickable(
        (
            By.CSS_SELECTOR,
            "button.o-button.primary"
        )
    )
)

driver.execute_script(
    "arguments[0].click();",
    final_login
)

print("CMA CGM Login Success ✅")

time.sleep(10)

driver.execute_script("window.location.href='/ebusiness/pricing/instant-Quoting'")

try:
    # wait until element is visible
    title_el = WebDriverWait(driver, 30).until(
        EC.visibility_of_element_located((By.ID, "pageTitle"))
    )

    # get text safely
    title = title_el.get_attribute("textContent").strip()

    print("TITLE:", repr(title))

    if "We are improving the eBusiness area" in title:
        print("Maintenance page detected")

        data.append([
            today,
            "CMACGM",
            origin_port_name,
            destination_port_name,
            cont_size,
            20000,
            "", "", "", "", "", "", "",
            "Site Under Maintenance"
        ])

        # create dataframe
        new_df = pd.DataFrame(data, columns=columns)

        try:
            # read old file
            old_df = pd.read_excel(file_path)

            # merge old + new
            final_df = pd.concat([old_df, new_df], ignore_index=True)

        except Exception as read_error:
            print("Excel read error:", read_error)

            # if file not found
            final_df = new_df

        # ALWAYS SAVE
        final_df.to_excel(file_path, index=False)

        print("Data saved successfully")

        driver.quit()
        sys.exit()

except:
    pass


time.sleep(10)
try:
    close = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[class='PopupSurvey_26859_1601074_CloseButton']")))
    close.click()
except:
    pass    
origin_input = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[aria-controls='sortedAutocompletePopup-origin-field']")))
origin_input.send_keys(origin_port_code)
origin_options = wait.until(EC.presence_of_all_elements_located(
    (By.CSS_SELECTOR, "li.place-suggestion")
))
driver.execute_script("arguments[0].click();", origin_options[0])

dest_input = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[aria-controls='sortedAutocompletePopup-destination-field']")))
dest_input.send_keys(destination_port_code)

dest_option = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[id='sortedAutocompletePopup-destination-field'] li.place-suggestion")))

driver.execute_script("arguments[0].click();", dest_option)
time.sleep(2)
try:
    # POL Click
    pol = driver.find_element(By.XPATH, "//label[text()='POL']/following::i[contains(@class,'el-select__caret')][1]")
    pol.click()
    time.sleep(3)
    # dropdown option
    first_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[@role='option']//b")))
    driver.execute_script("arguments[0].click();", first_option)  
    time.sleep(2)
except:    
    pass

cardss = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "ul.container-size-switches div.content")))
for card in cardss:
    con_size=card.find_element(By.CSS_SELECTOR, "span[class^='ico']").text
    if container_size in con_size:
        add_btn = card.find_element(By.CSS_SELECTOR, "button.add-button")
        driver.execute_script("arguments[0].click();", add_btn)
        time.sleep(2)
        weight_input = card.find_element(By.CSS_SELECTOR, "span[name='weightPerContainer'] input")
        time.sleep(2)
        weight_input.send_keys("20000")       
        break            
            
try:
    all=driver.find_element(By.CSS_SELECTOR, "div[class='el-select commodity-select rm-tab']")
    all.click()
    time.sleep(2)    
except:
    pass    
all_kinds = wait.until(EC.element_to_be_clickable(
    (By.XPATH, "//span[text()='Freight All Kinds']")
))
driver.execute_script("arguments[0].click();", all_kinds)
time.sleep(5)

try:
    driver.find_element(By.CSS_SELECTOR, "div[class='el-select multiselect-autocomplete']").click()
    time.sleep(2)
    bco = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='BCO']")))
    driver.execute_script("arguments[0].click();", bco)
except:
    pass    

submit_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[id='SearchQuote']")))
submit_btn.click()
time.sleep(10)

cards = []

try:

    no_data = driver.find_element(
        By.XPATH,
        "//h2[contains(text(),'We apologize')]"
    ).text.strip()

    print(no_data)

    if no_data == "We apologize,":

        data.append([
            today,
            "CMACGM",
            origin_port_name,
            destination_port_name,
            cont_size,
            20000,
            "", "", "", "", "", "", "",
            "No Data"
        ])

except:

    cards = wait.until(
        EC.presence_of_all_elements_located(
            (By.CSS_SELECTOR, "div.wrap-card")
        )
    )

print("Total cards:", len(cards))

# =========================
# LOOP CARDS card
# =========================
for lo, card in enumerate(cards):

    print("Card Number:", lo)

    try:

        # ================= DATE =================
        dates = card.find_elements(By.CSS_SELECTOR, "span.date")

        start_date = ""
        end_date = ""

        if len(dates) >= 2:

            try:
                start_date = dates[0].text.split(',')[1].strip()
                end_date = dates[1].text.split(',')[1].strip()
            except:
                pass

        # ================= DEFAULT VALUES =================
        ocean_freight = 0.0
        charges_freight = 0.0
        charges_import = 0.0
        charges_export = 0.0

        # ================= CLICK DETAILS =================
        try:

            details_btn = card.find_element(
                By.XPATH,
                ".//label[.//span[contains(text(),'Details')]]"
            )

            driver.execute_script(
                "arguments[0].click();",
                details_btn
            )

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "span.charges-detail")
                )
            )

            time.sleep(1.5)

            rows = driver.find_elements(
                By.CSS_SELECTOR,
                "div.el-table__body-wrapper tbody tr.el-table__row"
            )

            if not rows:

                print("No rows found after clicking details")
                continue

            for row in rows:

                try:

                    label = row.find_element(
                        By.CSS_SELECTOR,
                        "span.charges-detail"
                    ).text.strip()

                    value_text = row.find_element(
                        By.CSS_SELECTOR,
                        "td:nth-child(3) span"
                    ).text.strip()

                    try:
                        value = float(
                            value_text.replace(",", "").strip()
                        )
                    except:
                        value = 0.0

                    label_lower = label.lower()

                    # ================= MAPPING =================
                    if "ocean freight" in label_lower:
                        ocean_freight = value

                    elif "freight" in label_lower and "as per" in label_lower:
                        charges_freight = value

                    elif "import" in label_lower:
                        charges_import = value

                except:
                    continue

        except Exception as e:

            print("Details click error:", e)

        # ================= TOTAL =================
        total = round(
            ocean_freight +
            charges_freight +
            charges_export +
            charges_import,
            1
        )

        # ================= DEBUG =================
        print("Start Date:", start_date)
        print("End Date:", end_date)
        print("Ocean Freight:", ocean_freight)
        print("Charges Freight:", charges_freight)
        print("Charges Import:", charges_import)
        print("TOTAL:", total)

        # ================= SAVE =================
        data.append([
            today,
            "CMACGM",
            origin_port_name,
            destination_port_name,
            cont_size,
            20000,
            start_date,
            end_date,
            ocean_freight,
            charges_freight,
            charges_export,
            charges_import,
            total,
            ""
        ])

    except Exception as e:

        print("Card error:", e)
        continue
 
new_df = pd.DataFrame(data, columns=columns)
try:
    old_df = pd.read_excel(file_path)
    final_df = pd.concat([old_df, new_df], ignore_index=True)
except:
    final_df = new_df
final_df.to_excel(file_path, index=False)
print("Data saved successfully")
try:
    driver.quit()
except:
    pass