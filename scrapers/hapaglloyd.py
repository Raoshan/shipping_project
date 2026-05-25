from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time
import undetected_chromedriver as uc
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import pandas as pd    
import os
# ================= FILE =================
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

records = []


# ================= DRIVER =================
options = uc.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument("--disable-blink-features=AutomationControlled")
profile_path = os.path.join(BASE_DIR, "chrome", "hapag")
options.add_argument(f"--user-data-dir={profile_path}")


driver = uc.Chrome(
    options=options,
    version_main=148,
    use_subprocess=True
)

driver.maximize_window()

wait = WebDriverWait(driver, 30)

# ================= OPEN SITE =================

driver.get("https://www.hapag-lloyd.com/en/home.html")

time.sleep(10)

# ================= ACCEPT COOKIES =================

try:

    wait.until(
        EC.element_to_be_clickable(
            (By.ID, "accept-recommended-btn-handler")
        )
    ).click()

    print("cookies accepted")

except:
    pass

time.sleep(5)

# ================= LOGIN CHECK =================

body = driver.find_element(
    By.TAG_NAME,
    "body"
).text.lower()

if (
    "gitanjali" in body
    or "my hapag" in body
    or "sign out" in body
    or "logout" in body
    or "tanu impex" in body
):

    print("already logged in")

else:

    print("login required") #div.hal-navigation-top-login-web

    try:

        # ================= CLICK LOGIN =================

        login_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div.hal-navigation-top-login-web"))).click()  

        print("clicked login")
        time.sleep(10)

        # ================= USERNAME =================

        username = wait.until(
            EC.visibility_of_element_located(
                (By.ID, "signInName")
            )
        )

        username.clear()

        username.send_keys(
            "shipping@tanuimpex.com"
        )

        print("username entered")

        # ================= PASSWORD =================

        password = wait.until(
            EC.visibility_of_element_located(
                (By.ID, "password")
            )
        )

        password.clear()

        password.send_keys(
            "MyHapag123"
        )

        print("password entered")

        time.sleep(2)

        # ================= LOGIN BUTTON =================

        driver.find_element(
            By.ID,
            "next"
        ).click()

        print("login submitted")

        time.sleep(20)

        # ================= FINAL CHECK =================

        body = driver.find_element(
            By.TAG_NAME,
            "body"
        ).text.lower()

        if (
            "gitanjali" in body
            or "my hapag" in body
            or "sign out" in body
        ):

            print("login completed")

        else:

            print("login may have failed")

    except Exception as e:

        print("login failed")
        print(e)

# ================= READY =================

print("ready for scraping")

driver.execute_script("window.location.href='/solutions/new-quote/#/simple'")

time.sleep(20)
import sys
# ================= INPUT =================
today = datetime.today().strftime('%d-%b-%Y')

origin_port_code = sys.argv[1]
origin_port_name = sys.argv[2]
destination_port_code = sys.argv[3]
destination_port_name = sys.argv[4]
container_type = sys.argv[5]
print("Received arguments:", origin_port_code, origin_port_name, destination_port_code, destination_port_name, container_type)
container_map = {
    "20' General Purpose": "1 × 20' ST",
    "40' General Purpose": "1 × 40' ST",
    "40' General Purpose High Cube": "1 × 40' HC"
}

size_map = {
    "20' General Purpose": "20STD", 
    "40' General Purpose": "40STD", 
    "40' General Purpose High Cube": "40HC" 
}
short_name = size_map.get(container_type.strip()) 

if not short_name:
    raise Exception("Container mapping not found")

cont_size = container_map.get(container_type, "")

try:
    close_btn = driver.find_element(
        By.XPATH,
        '//button[contains(@class,"q-dialog__x")]'
    )

    close_btn.click()
except:
    pass
time.sleep(2)
# Origin
origin_input = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[aria-label='Start Location (Routing 1)']")))
origin_input.send_keys(origin_port_code)
time.sleep(2)

wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div.q-item__section--main"))).click()
time.sleep(2)

# Destination
dest_input = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[aria-label='End Location (Routing 1)']")))
dest_input.send_keys(destination_port_code)
time.sleep(2)
wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div.q-item__section--main"))).click()
time.sleep(2)

# Container
container_dropdown = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div[data-testid='container-input']")))
driver.execute_script("arguments[0].click();", container_dropdown)
time.sleep(2)
containers = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "span.typography-apps-label-default")))
time.sleep(2)

for cont in containers:
    if container_type in cont.text:
        driver.execute_script("arguments[0].click();", cont)
        break

# Submit
wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit']"))).click()
time.sleep(20)

try:
    pop_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='Close']]")))
    pop_btn.click()
except:
    pass

def get_column_index(table, target):

    headers = table.find_elements(By.CSS_SELECTOR, "thead th")

    for i, h in enumerate(headers):

        txt = h.text.strip().upper().replace(" ", "")

        print("HEADER =", txt)

        if target.upper().replace(" ", "") == txt:
            print("MATCHED COLUMN =", i)
            return i

    return None


# =========================
# Clean Amount
# =========================

import re

def clean_int(text):

    try:

        text = (
            str(text)
            .replace(",", "")
            .replace("\u202f", "")
            .replace("\xa0", "")
            .replace("$", "")
            .replace("USD", "")
            .replace("EUR", "")
            .replace("INR", "")
            .strip()
        )

        # extract number
        match = re.search(r"[-+]?\d*\.?\d+", text)

        if not match:
            return 0

        return float(match.group())

    except Exception as e:

        print("Conversion error:", text, e)

        return 0


# =========================
# Extract Final Data
# =========================

def extract_final_data(short_name):

    modal = wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, ".q-dialog")
        )
    )

    result = {
        "ocean": 0,
        "freight_surcharge": 0,
        "export": 0,
        "import": 0
    }

    tables = modal.find_elements(
        By.CSS_SELECTOR,
        "table.q-table"
    )

    print("TOTAL TABLES =", len(tables))

    for table in tables:

        try:

            parent_text = table.find_element(
                By.XPATH,
                "./ancestor::div[contains(@class,'q-table__container')]"
            ).text.upper()

        except:
            parent_text = table.text.upper()

        print("\n======================")
        print(parent_text[:200])

        # dynamic column index
        col_index = get_column_index(table, short_name)

        print("COL INDEX =", col_index)

        rows = table.find_elements(
            By.XPATH,
            ".//tbody/tr"
        )

        # =========================
        # FREIGHT CHARGES
        # =========================

        if (
            "FREIGHT CHARGES" in parent_text
            and "SURCHARGES" not in parent_text
        ):

            for row in rows:

                try:

                    tds = row.find_elements(By.TAG_NAME, "td")

                    print("FREIGHT TD =", len(tds))

                    if col_index is not None and len(tds) > col_index:

                        value = clean_int(
                            tds[col_index].text
                        )

                    else:

                        value = clean_int(
                            tds[-1].text
                        )

                    result["ocean"] += value

                except Exception as e:
                    print("Ocean Error =", e)

        # =========================
        # FREIGHT SURCHARGES
        # =========================

        elif "FREIGHT SURCHARGES" in parent_text:

            for row in rows:

                try:

                    tds = row.find_elements(By.TAG_NAME, "td")

                    print("SURCHARGE TD =", len(tds))

                    if col_index is not None and len(tds) > col_index:

                        value = clean_int(
                            tds[col_index].text
                        )

                    else:

                        value = clean_int(
                            tds[-1].text
                        )

                    result["freight_surcharge"] += value

                except Exception as e:
                    print("Freight Surcharge Error =", e)

        # =========================
        # EXPORT SURCHARGES
        # =========================

        elif "EXPORT SURCHARGES" in parent_text:

            for row in rows:

                try:

                    tds = row.find_elements(By.TAG_NAME, "td")

                    print("EXPORT TD =", len(tds))

                    if col_index is not None and len(tds) > col_index:

                        value = clean_int(
                            tds[col_index].text
                        )

                    else:

                        value = clean_int(
                            tds[-1].text
                        )

                    result["export"] += value

                except Exception as e:
                    print("Export Error =", e)

        # =========================
        # IMPORT SURCHARGES 94
        # =========================

        elif "IMPORT SURCHARGES" in parent_text:

            for row in rows:

                try:

                    tds = row.find_elements(By.TAG_NAME, "td")

                    print("IMPORT TD =", len(tds))

                    if col_index is not None and len(tds) > col_index:

                        value = clean_int(
                            tds[col_index].text
                        )

                    else:

                        value = clean_int(
                            tds[-1].text
                        )

                    result["import"] += value

                except Exception as e:
                    print("Import Error =", e)

    return result 

def format_date(date_str):
    try:
        try:
            dt = datetime.strptime(date_str.strip(), "%d %b %Y")  # Apr
        except:
            dt = datetime.strptime(date_str.strip(), "%d %B %Y")  # April

        return dt.strftime("%d-%b-%Y")   # ✅ 05-May-2026
    except:
        print("Date format issue:", date_str)
        return ""

# ================= MAIN LOOP =================
cards = []
try:
    cards = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "button.carousel__item")))
except:
    records.append([today, "HAPAG", origin_port_name, destination_port_name, cont_size, 20000, "", "", "", "", "", "", "", "No Data"])

print("Total cards:", len(cards))


try:

    close_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.q-dialog__x")))
    driver.execute_script("arguments[0].click();", close_btn)
except:
    try:

        close_btn2 = wait.until(EC.element_to_be_clickable((By.XPATH, '/button[.//span[text()="Close"]]')))
        driver.execute_script("arguments[0].click();", close_btn2)
    except:
        pass

cards_count = len(driver.find_elements(By.CSS_SELECTOR, "button.carousel__item")) #Import row error
print("Total Cards =", cards_count)
for i in range(cards_count):

    # refresh cards
    cards = wait.until(
        EC.presence_of_all_elements_located(
            (By.CSS_SELECTOR, "button.carousel__item")
        )
    )

    # click card
    driver.execute_script("arguments[0].click();", cards[i])

    time.sleep(3)

    # all Price Breakdown buttons
    buttons = driver.find_elements(
        By.XPATH,
        "//button[.//span[text()='Price Breakdown'] and not(@disabled)]"
    )

    print("Total Price Breakdown Buttons =", len(buttons))

    # loop all buttons
    for j in range(len(buttons)):

        try:
            # refresh buttons again
            buttons = driver.find_elements(
                By.XPATH,
                "//button[.//span[text()='Price Breakdown'] and not(@disabled)]"
            )

            btn = buttons[j]  
            quote_type = "Quick Quotes" if j == 0 else "Quick Quotes Spot"            
            time.sleep(1)
            driver.execute_script(
                "arguments[0].click();",
                btn
            )

            print(f"Clicked Button {j+1}")

            time.sleep(3)

            # ===== YOUR DATA EXTRACTION =====
            validity = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.offer-information__quote--validity")))
            items = validity.find_elements(By.CSS_SELECTOR, "div.hal-data-item")
            valid_from = ""
            valid_to = ""
            for item in items:
                label = item.find_element(By.CSS_SELECTOR, ".hal-data-item__label").text.strip()
                value = item.find_element(By.CSS_SELECTOR, ".hal-data-item__content").text.strip()
                if "Valid from" in label:
                   valid_from = value
                elif "Valid to" in label:
                    valid_to = value
                elif "ETD" in label:
                    valid_from = value    
                elif "ETA" in label:
                    valid_to = value    
            start_date = format_date(valid_from)
            end_date = format_date(valid_to)
            print(start_date)
            print(end_date)
            result = extract_final_data(short_name)
            print("Ocean =", result["ocean"])
            print("FSurcharge =", result["freight_surcharge"])
            print("export =", result["export"])
            result["import"] = round(result["import"] / 96, 1)
            print("import =", result["import"])
            total = round(result["ocean"] + result["freight_surcharge"] + result["export"]+result["import"], 1)
            records.append([today,
                "HAPAG",
                origin_port_name,
                destination_port_name,
                cont_size,
                20000,
                start_date,
                end_date,
                result["ocean"],
                result["freight_surcharge"],
                result["export"],
                result["import"],
                total,
                quote_type
            ])
            # close modal
            close_btn = wait.until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "button.q-dialog__x")
                )
            )
            driver.execute_script(
                "arguments[0].click();",
                close_btn
            )
            time.sleep(2)

        except Exception as e:
            print("Button Error =", e)

# ================= SAVE =================
new_df = pd.DataFrame(records, columns=columns)

try:
    old_df = pd.read_excel(file_path)
    final_df = pd.concat([old_df, new_df], ignore_index=True)
except:
    final_df = new_df

final_df.to_excel(file_path, index=False)

print("Data saved successfully")
driver.quit()
