import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from datetime import datetime
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
from selenium.webdriver.common.keys import Keys
import sys
import os
from selenium.common.exceptions import *

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

# =========================================================
# DRIVER SETUP
# =========================================================
options = uc.ChromeOptions()

options.add_argument("--headless=new")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--disable-gpu")

driver = uc.Chrome(
    options=options,
    browser_executable_path="/usr/bin/chromium",
    driver_executable_path="/usr/bin/chromedriver",
    use_subprocess=True
)
wait = WebDriverWait(driver, 30)

# =========================================================
# OPEN WEBSITE
# =========================================================
driver.get("https://www.mymsc.com")
driver.maximize_window()
time.sleep(8)

# =========================================================
# ACCEPT COOKIES
# =========================================================
try:

    allow_btn = driver.find_element(
        By.CSS_SELECTOR,
        "button#onetrust-accept-btn-handler"
    )

    allow_btn.click()

    print("Cookies Accepted")

    time.sleep(2)

except:

    print("Cookies Popup Not Found")


# =========================================================
# EMAIL
# =========================================================

email_input = wait.until(
    EC.presence_of_element_located((
        By.CSS_SELECTOR,
        "input[type='email']"
    ))
)

email_input.clear()

email_input.send_keys(
    "sales@tanuimpex.com"
)

print("Email Entered")

time.sleep(2)


# =========================================================
# NEXT BUTTON
# =========================================================

next_btn = wait.until(
    EC.element_to_be_clickable((
        By.CSS_SELECTOR,
        "button[type='button']"
    ))
)

next_btn.click()

print("Next Button Clicked")

time.sleep(5)


# =========================================================
# PASSWORD
# =========================================================

password_input = wait.until(
    EC.presence_of_element_located((
        By.CSS_SELECTOR,
        "input[type='password']"
    ))
)

password_input.clear()

password_input.send_keys(
    "E@Lzy$-Sij82V2*"
)

print("Password Entered")

time.sleep(2)


# =========================================================
# LOGIN BUTTON
# =========================================================

login_btn = wait.until(
    EC.element_to_be_clickable((
        By.CSS_SELECTOR,
        "button#next"
    ))
)

login_btn.click()

print("LOGIN COMPLETED")
time.sleep(15)

print("READY FOR SCRAPING")

driver.execute_script("window.location.href='/instantquote'")
time.sleep(10)

origin_port_code = sys.argv[1]
origin_port_name = sys.argv[2]
destination_port_code = sys.argv[3]
destination_port_name = sys.argv[4]
container_size = sys.argv[5]
print("Received arguments:", origin_port_code, origin_port_name, destination_port_code, destination_port_name, container_size)

container_map = {
    "20DV": "1 × 20' ST",
    "40DV": "1 × 40' ST",
    "40HC": "1 × 40' HC"
}

# step 1: main component pakdo
shadow_host = driver.find_element(By.CSS_SELECTOR, "mymsc-instantquote-app")

# step 2: shadow root access karo
shadow_root = driver.execute_script("return arguments[0].shadowRoot", shadow_host)

# step 3: ab andar search karo
all_cont_size = shadow_root.find_elements(
    By.CSS_SELECTOR, "div[data-test-id^='equipment-sizetype-container']"
)

print("Total Containers:", len(all_cont_size))

# 🔥 LOOP (no stale issue)
for i in range(len(all_cont_size)):   # total 3 containers
    try:
        # ✅ har baar fresh elements lo
        all_cont_size = shadow_root.find_elements(
            By.CSS_SELECTOR, "div[data-test-id^='equipment-sizetype-container']"
        )

        con_size = all_cont_size[i]

        actual_size = con_size.find_element(
            By.CSS_SELECTOR, "span.MuiTypography-caption"
        ).text.strip()

        checkbox = con_size.find_element(By.CSS_SELECTOR, "input[type='checkbox']")
        label = con_size.find_element(By.CSS_SELECTOR, "label")

        print("Found:", actual_size)

        # ✅ Select required
        if actual_size == container_size:
            if not checkbox.is_selected():
                driver.execute_script("arguments[0].click();", label)
                print("✔ Selected:", actual_size)

            # ✅ SAVE VALUE
            final_value = container_map.get(actual_size, "")
            print("Saved Value:", final_value)

        # ❌ Uncheck others
        else:
            if checkbox.is_selected():
                driver.execute_script("arguments[0].click();", label)
                print("✖ Unchecked:", actual_size)

    except Exception as e:
        print("Retrying due to:", e)

# ✅ Final output
print("Final Container Value:", container_map.get(container_size, ""))
cont_size=container_map.get(container_size, "")

driver.execute_script(f"""
let host = document.querySelector('mymsc-instantquote-app');
let root = host.shadowRoot;
let input = root.querySelector('#origin');

input.focus();

// React setter
let setter = Object.getOwnPropertyDescriptor(
    window.HTMLInputElement.prototype, "value"
).set;

setter.call(input, "{origin_port_code}");

// events trigger
input.dispatchEvent(new Event('input', {{ bubbles: true }}));
input.dispatchEvent(new Event('change', {{ bubbles: true }}));
""")

time.sleep(5)
# STEP 2: CLICK OPTION

driver.execute_script(f"""
let host = document.querySelector('mymsc-instantquote-app');
let root = host.shadowRoot;

let options = root.querySelectorAll('li');

for (let opt of options) {{
    if (opt.innerText.includes("{origin_port_code}")) {{
        opt.click();
        break;
    }}
}}
""")
time.sleep(5)
print("USNYC SELECT HO GAYA")

driver.execute_script(f"""
let host = document.querySelector('mymsc-instantquote-app');
let root = host.shadowRoot;
let input = root.querySelector('#destination');

input.focus();

// React setter
let setter = Object.getOwnPropertyDescriptor(
    window.HTMLInputElement.prototype, "value"
).set;

setter.call(input, "{destination_port_code}");

// events trigger
input.dispatchEvent(new Event('input', {{ bubbles: true }}));
input.dispatchEvent(new Event('change', {{ bubbles: true }}));
""")

time.sleep(5)

# STEP 2: SELECT OPTION (KEYBOARD)

driver.execute_script(f"""
let host = document.querySelector('mymsc-instantquote-app');
let root = host.shadowRoot;
let input = root.querySelector('#destination');

input.focus();
input.dispatchEvent(new KeyboardEvent('keydown', {{key: 'ArrowDown', bubbles: true}}));
input.dispatchEvent(new KeyboardEvent('keydown', {{key: 'Enter', bubbles: true}}));
""")

print(f"DESTINATION SELECT HO GAYA: {destination_port_code}")

driver.execute_script("""
let host = document.querySelector('mymsc-instantquote-app');
let root = host.shadowRoot;

// button select
let btn = root.querySelector('#search-rate-button');

btn.click();
""")
print("Search button clicked")
time.sleep(20)

today = datetime.today().strftime('%d-%b-%Y')

# =====================================
# SHADOW ROOT
# =====================================

def get_root():

    host = driver.find_element(
        By.CSS_SELECTOR,
        "mymsc-instantquote-app"
    )

    return driver.execute_script(
        "return arguments[0].shadowRoot",
        host
    )


# =====================================
# GET ALL UPPER CARDS
# =====================================

def get_upper_cards():

    root = get_root()

    return root.find_elements(
        By.CSS_SELECTOR,
        "div.carousel-card-container"
    )


# =====================================
# GET RATE CARDS
# =====================================

def get_rate_cards():

    root = get_root()

    return root.find_elements(
        By.CSS_SELECTOR,
        "div.rateCardBox"
    )


# =====================================
# CLOSE PDF
# =====================================

def close_pdf():

    try:

        root = get_root()

        close_btn = root.find_element(
            By.CSS_SELECTOR,
            "button.myMSC-icon-close.close-btn"
        )

        driver.execute_script(
            "arguments[0].click();",
            close_btn
        )

        print("PDF CLOSED")

        time.sleep(2)

    except Exception as e:

        print("PDF CLOSE ERROR:", e)


# =====================================
# CLICK UPPER CARD PROPERLY
# =====================================

def click_upper_card(index):

    upper_cards = get_upper_cards()

    card = upper_cards[index]

    # scroll into view
    driver.execute_script("""
        arguments[0].scrollIntoView({
            behavior:'smooth',
            block:'center',
            inline:'center'
        });
    """, card)

    time.sleep(2)

    # actual clickable card
    clickable = card.find_element(
        By.CSS_SELECTOR,
        ".search-result-carousel-card"
    )

    # remove active class info
    print(
        "ACTIVE BEFORE:",
        "selected" in clickable.get_attribute("class")
    )

    # click using JS
    driver.execute_script(
        "arguments[0].click();",
        clickable
    )

    time.sleep(4)

    # verify active changed
    upper_cards = get_upper_cards()

    active_card = upper_cards[index].find_element(
        By.CSS_SELECTOR,
        ".search-result-carousel-card"
    )

    is_selected = "selected" in active_card.get_attribute("class")

    print(f"UPPER CARD {index+1} CLICKED")
    print("ACTIVE AFTER:", is_selected)

    return is_selected


# =====================================
# MAIN PROCESS
# =====================================

total_upper_cards = len(get_upper_cards())

print("TOTAL UPPER CARDS:", total_upper_cards)


# =====================================
# UPPER LOOP
# =====================================

from selenium.webdriver.common.by import By
from selenium.common.exceptions import *
from datetime import datetime
import time


# =====================================
# MAIN LOOP
# =====================================

for i in range(total_upper_cards):

    print("\n==============================")
    print(f"UPPER CARD {i+1}")
    print("==============================")

    try:

        # =====================================
        # CLICK CURRENT UPPER CARD
        # =====================================

        active = click_upper_card(i)

        if not active:
            print("CARD NOT ACTIVATED")
            continue

        time.sleep(3)

        # =====================================
        # GET RATE CARDS
        # =====================================

        total_rate_cards = len(get_rate_cards())

        print("RATE CARDS:", total_rate_cards)

        # =====================================
        # RATE CARD LOOP
        # =====================================

        for r in range(total_rate_cards):

            print(f"\nRATE CARD {r+1}")

            try:

                # fresh rate cards
                rate_cards = get_rate_cards()

                rate_card = rate_cards[r]

                # =====================================
                # SHIPPING WINDOW
                # =====================================

                shipping_window = rate_card.find_element(
                    By.CSS_SELECTOR,
                    "[data-test-id^='ShippingWindow_'] b"
                ).text

                start_raw, end_raw = shipping_window.split(" - ")

                start_date = datetime.strptime(
                    start_raw.strip(),
                    "%d %b %Y"
                ).strftime("%d-%b-%Y")

                end_date = datetime.strptime(
                    end_raw.strip(),
                    "%d %b %Y"
                ).strftime("%d-%b-%Y")

                print("START DATE:", start_date)
                print("END DATE:", end_date)

                # =====================================
                # SHOW DETAILS BUTTONS
                # =====================================

                detail_buttons = rate_card.find_elements(
                    By.CSS_SELECTOR,
                    "[data-test-id*='showDetailsIcon']"
                )

                print("DETAIL BUTTONS:", len(detail_buttons))

                # =====================================
                # DETAIL LOOP
                # =====================================

                for d in range(len(detail_buttons)):

                    try:

                        print(f"Opening Detail {d+1}")

                        # fresh elements again
                        rate_cards = get_rate_cards()

                        rate_card = rate_cards[r]

                        detail_buttons = rate_card.find_elements(
                            By.CSS_SELECTOR,
                            "[data-test-id*='showDetailsIcon']"
                        )

                        detail_btn = detail_buttons[d]

                        # scroll
                        driver.execute_script("""
                            arguments[0].scrollIntoView({
                                behavior:'smooth',
                                block:'center'
                            });
                        """, detail_btn)

                        time.sleep(1)

                        # click detail
                        driver.execute_script(
                            "arguments[0].click();",
                            detail_btn
                        )

                        print(f"DETAIL {d+1} OPENED")

                        time.sleep(4)

                        # =====================================
                        # EXTRACT DATA
                        # =====================================

                        records = driver.execute_script("""

                        let result = {
                            freight_charge: 0,
                            freight_surcharges: 0,
                            export_surcharges: 0,
                            import_surcharges: 0
                        };

                        let root = document.querySelector(
                            'mymsc-instantquote-app'
                        ).shadowRoot;

                        let modal = root.querySelector(
                            '[data-test-id="BreakdownModal"]'
                        );

                        let rows = modal.querySelectorAll(
                            '.standard-charges-row'
                        );

                        let section = "";

                        let import_inr = 0;
                        let import_usd = 0;

                        rows.forEach(row => {

                            let header = row.querySelector("strong");

                            if(header){

                                let text = header.innerText.trim();

                                if(text.includes("Freight Charge"))
                                    section = "freight_charge";

                                else if(text.includes("Freight Surcharges"))
                                    section = "freight_surcharges";

                                else if(text.includes("Export Surcharges"))
                                    section = "export_surcharges";

                                else if(text.includes("Import Surcharges"))
                                    section = "import_surcharges";
                            }

                            let amounts = row.querySelectorAll(
                                ".amount-field strong"
                            );

                            amounts.forEach(el => {

                                let txt = el.innerText.trim();

                                let num = txt.replace(
                                    /[^0-9.]/g,
                                    ""
                                );

                                if(!num) return;

                                let val = parseFloat(num);

                                if(section === "import_surcharges"){

                                    if(txt.includes("INR"))
                                        import_inr += val;

                                    else
                                        import_usd += val;
                                }

                                else if(section === "freight_charge"){
                                    result.freight_charge += val;
                                }

                                else if(section === "freight_surcharges"){
                                    result.freight_surcharges += val;
                                }

                                else if(section === "export_surcharges"){
                                    result.export_surcharges += val;
                                }

                            });

                        });

                        let import_final =
                            (import_inr / 96) + import_usd;

                        result.import_surcharges =
                            parseFloat(import_final.toFixed(1));

                        return result;

                        """)

                        # =====================================
                        # TOTAL
                        # =====================================
                        total = round(sum(records.values()), 1)

                        records["freight_charge"] = round(
                            records["freight_charge"], 1
                        )

                        records["freight_surcharges"] = round(
                            records["freight_surcharges"], 1
                        )

                        records["export_surcharges"] = round(
                            records["export_surcharges"], 1
                        )

                        records["import_surcharges"] = round(
                            records["import_surcharges"], 1
                        )

                        print("Freight:", records["freight_charge"])
                        print("Surcharges:", records["freight_surcharges"])
                        print("Export:", records["export_surcharges"])
                        print("Import:", records["import_surcharges"])
                        print("TOTAL:", total)

                        # =====================================
                        # SAVE DATA
                        # =====================================
                        data.append([
                            today,
                            "MSC",
                            origin_port_name,
                            destination_port_name,
                            cont_size,
                            20000,
                            start_date,
                            end_date,
                            records["freight_charge"],
                            records["freight_surcharges"],
                            records["export_surcharges"],
                            records["import_surcharges"],
                            total,
                            ""
                        ])

                        # =====================================
                        # CLOSE PDF
                        # =====================================
                        close_pdf()

                        time.sleep(2)

                    except Exception as e:

                        print(f"DETAIL {d+1} ERROR:", e)

                        try:
                            close_pdf()
                        except:
                            pass

                        continue

            except Exception as e:
                print("RATE CARD ERROR:", e)

                try:
                    close_pdf()
                except:
                    pass

                continue

    except Exception as e:
        print("UPPER CARD ERROR:", e)
        continue


# =====================================
# NO DATA
# =====================================
if total_upper_cards == 0:
    data.append([today, "MSC", origin_port_name, destination_port_name, cont_size, 20000, "", "", "", "", "", "", "", "No Data"])
    print("No cards found")

print("ALL COMPLETED")  
        
new_df = pd.DataFrame(data, columns=columns)
try:
    old_df = pd.read_excel(file_path)
    final_df = pd.concat([old_df, new_df], ignore_index=True)
except:
    final_df = new_df
final_df.to_excel(file_path, index=False)

print("Data saved successfully")    

driver.quit()