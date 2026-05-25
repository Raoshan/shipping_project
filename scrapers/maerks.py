import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
import time
import os
# ================= FILE =================
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))            
file_path = os.path.join(BASE_DIR, "global", "rates.xlsx")

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

# -------- FIREFOX SETUP --------
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

driver.get("https://www.maersk.com/")
driver.maximize_window()
time.sleep(8)

try:
    driver.find_element(By.CSS_SELECTOR, "button[data-test='coi-allow-all-button']").click()
    time.sleep(2)
except:
    pass
try:    
    driver.find_element(By.CSS_SELECTOR, "a[class='ign-header__primary__actions__links__link ign-button-primary   ign-track']").click()
except:
    pass
time.sleep(10) 	

# thoda wait
# ========= CHECK LOGIN =========
already_logged_in = False

try:
    # Agar ye naam mil gaya -> login already hai
    driver.find_element(
        By.XPATH,
        "//*[contains(text(),'GITANJALI AGGARWAL')]"
    )

    already_logged_in = True
    print("Already Logged In ✅")

except:
    print("Login Required 🔐")


# ========= LOGIN ONLY IF NEEDED =========
if not already_logged_in:

    for el in driver.find_elements(By.CSS_SELECTOR, "mc-input"):

        shadow = driver.execute_script(
            "return arguments[0].shadowRoot",
            el
        )

        inp = shadow.find_element(By.CSS_SELECTOR, "input")

        if inp.get_attribute("id") == "mc-input-username":
            inp.send_keys("Tanuimpex")

        elif inp.get_attribute("id") == "mc-input-password":
            inp.send_keys("MyMaersk123")

    time.sleep(2)

    driver.execute_script("""
    let btn = document.querySelector("mc-button#login-submit-button");
    btn.shadowRoot.querySelector("button").click();
    """)

    print("Login Done ✅")

time.sleep(20)

try:
    prices_btn = driver.find_element(By.XPATH, "//a[contains(text(),'Prices')]")
    driver.execute_script("arguments[0].click();", prices_btn)
    time.sleep(10)
    instant_btn = driver.find_element(By.XPATH, "//a[contains(text(),'Instant prices')]")
    driver.execute_script("arguments[0].click();", instant_btn)
    time.sleep(10)
except:
    driver.get("https://www.maersk.com/book/")
    time.sleep(10)    
    

# Continue
try:
    driver.execute_script("""document.querySelector("mc-button#login-submit-button").shadowRoot.querySelector("button").click();""")
    time.sleep(5)
except:
    pass    
time.sleep(10)

today = datetime.today().strftime('%d-%b-%Y')
import sys

origin_port_code = sys.argv[1]
origin_port_name = sys.argv[2]
destination_port_code = sys.argv[3]
destination_port_name = sys.argv[4]
container_type = sys.argv[5]
print("Received arguments:", origin_port_code, origin_port_name, destination_port_code, destination_port_name, container_type)
container_map = {
    "20 Dry Standard": "1 × 20' ST",
    "40 Dry Standard": "1 × 40' ST",
    "40 Dry High": "1 × 40' HC",
    "45 Dry High": "1 × 45' HC"
}

cont_size = container_map.get(container_type, "")
print("Mapped container type:", cont_size)

data = []

origin = driver.execute_script("""

const originText = arguments[0];

function sleep(ms){
    return new Promise(r => setTimeout(r, ms));
}

async function run(){

    // =========================
    // DEEP QUERY
    // =========================

    function deepQuery(selector){

        function find(root){

            let el = root.querySelector(selector);

            if(el) return el;

            let all = root.querySelectorAll("*");

            for(let i=0;i<all.length;i++){

                if(all[i].shadowRoot){

                    let found = find(all[i].shadowRoot);

                    if(found) return found;
                }
            }

            return null;
        }

        return find(document);
    }

    // =========================
    // DEEP QUERY ALL
    // =========================

    function deepQueryAll(selector){

        let results = [];

        function find(root){

            let els = root.querySelectorAll(selector);

            results.push(...els);

            let all = root.querySelectorAll("*");

            for(let el of all){

                if(el.shadowRoot){
                    find(el.shadowRoot);
                }
            }
        }

        find(document);

        return results;
    }

    // =========================
    // FIND INPUT
    // =========================

    let input = deepQuery("#mc-input-origin");

    if(!input)
        return "❌ ORIGIN INPUT NOT FOUND";

    input.focus();
    input.click();

    input.value = "";

    // =========================
    // REAL TYPING
    // =========================

    for(let ch of originText){

        input.value += ch;

        input.dispatchEvent(new InputEvent('input', {
            bubbles: true,
            data: ch,
            inputType: 'insertText'
        }));

        await sleep(50);
    }

    // =========================
    // WAIT DROPDOWN
    // =========================

    await sleep(2000);

    // =========================
    // FIND OPTIONS
    // =========================

    let options = deepQueryAll("mc-option");

    if(options.length === 0){
        return "❌ NO OPTIONS FOUND";
    }

    // =========================
    // CLICK MATCH OPTION
    // =========================

    for(let opt of options){

        let txt =
            (
                opt.innerText ||
                opt.textContent ||
                opt.getAttribute("label") ||
                ""
            ).trim().toLowerCase();

        if(txt.includes(originText.toLowerCase())){

            if(opt.shadowRoot){

                let btn =
                    opt.shadowRoot.querySelector("button");

                if(btn){

                    btn.click();

                    return "✅ ORIGIN SELECTED";
                }
            }

            opt.click();

            return "✅ FALLBACK CLICK";
        }
    }

    // =========================
    // FIRST OPTION FALLBACK
    // =========================

    options[0].click();

    return "✅ FIRST OPTION SELECTED";
}

return run();

""", origin_port_code)

time.sleep(5)

destination = driver.execute_script("""

const destinationText = arguments[0];

async function sleep(ms){
    return new Promise(r => setTimeout(r, ms));
}

async function run(){

    // =====================================
    // DEEP QUERY
    // =====================================

    function deepQuery(selector){

        function search(root){

            try{

                let el = root.querySelector(selector);

                if(el) return el;

                let all = root.querySelectorAll("*");

                for(let node of all){

                    if(node.shadowRoot){

                        let found = search(node.shadowRoot);

                        if(found) return found;
                    }
                }

            }catch(e){}

            return null;
        }

        return search(document);
    }

    // =====================================
    // DEEP QUERY ALL
    // =====================================

    function deepQueryAll(selector){

        let results = [];

        function search(root){

            try{
                results.push(...root.querySelectorAll(selector));
            }catch(e){}

            let all = [];

            try{
                all = root.querySelectorAll("*");
            }catch(e){}

            for(let el of all){

                if(el.shadowRoot){
                    search(el.shadowRoot);
                }
            }
        }

        search(document);

        return results;
    }

    // =====================================
    // INPUT
    // =====================================

    let input = deepQuery("#mc-input-destination");

    if(!input)
        return "❌ INPUT NOT FOUND";

    input.focus();
    input.click();

    await sleep(1000);

    // =====================================
    // CLEAR
    // =====================================

    input.value = "";

    input.dispatchEvent(new InputEvent("input",{
        bubbles:true,
        composed:true,
        data:"",
        inputType:"deleteContentBackward"
    }));

    await sleep(500);

    // =====================================
    // TYPE TEXT
    // =====================================

    for(let ch of destinationText){

        input.value += ch;

        input.dispatchEvent(new InputEvent("input",{
            bubbles:true,
            composed:true,
            data:ch,
            inputType:"insertText"
        }));

        await sleep(80);
    }

    await sleep(3000);

    // =====================================
    // OPTIONS
    // =====================================

    let options = deepQueryAll("mc-option");

    if(options.length === 0){

        return "❌ NO OPTIONS FOUND";
    }

    // =====================================
    // MATCH OPTION
    // =====================================

    let matched = null;

    let searchKey = destinationText
        .split(",")[0]
        .trim()
        .toLowerCase();

    for(let opt of options){

        let txt = (
            opt.innerText ||
            opt.textContent ||
            opt.getAttribute("label") ||
            ""
        )
        .replace(/\\s+/g," ")
        .trim()
        .toLowerCase();

        if(txt.includes(searchKey)){

            matched = opt;
            break;
        }
    }

    // fallback
    if(!matched){

        matched = options[0];
    }

    // =====================================
    // CLICK OPTION
    // =====================================

    matched.scrollIntoView({
        behavior:"smooth",
        block:"center"
    });

    await sleep(1000);

    if(matched.shadowRoot){

        let btn = matched.shadowRoot.querySelector("button");

        if(btn){

            btn.click();

        }else{

            matched.click();
        }

    }else{

        matched.click();
    }

    await sleep(3000);

    return {
        status : "✅ DESTINATION SELECTED",
        selectedValue : input.value,
        optionsFound : options.length
    };
}

return run();

""", destination_port_code)

time.sleep(5)

commodity = driver.execute_script("""
async function sleep(ms){
    return new Promise(r => setTimeout(r, ms));
}

// ================= DEEP QUERY =================

function deepAll(selector, root=document){

    let results = [];

    results.push(...root.querySelectorAll(selector));

    let all = root.querySelectorAll("*");

    for(let el of all){

        if(el.shadowRoot){

            results.push(...deepAll(selector, el.shadowRoot));
        }
    }

    return results;
}

// ================= FIND INPUT =================

let input = deepAll('input')
    .find(el =>
        (el.placeholder || '').toLowerCase().includes('type')
    );

if(!input){
    return "INPUT NOT FOUND";
}

input.scrollIntoView({block:'center'});
input.focus();
input.click();

await sleep(1000);

// clear
input.value = '';

let text = "metal coils";

// REAL typing
for(let ch of text){

    input.value += ch;

    input.dispatchEvent(new InputEvent('input',{
        bubbles:true,
        composed:true,
        data: ch,
        inputType:'insertText'
    }));

    await sleep(150);
}

await sleep(4000);

// ================= FIND MC-OPTION =================

let options = deepAll('mc-option');

if(options.length === 0){
    return "NO OPTIONS FOUND";
}

let target = null;

for(let opt of options){

    let txt = (opt.innerText || opt.textContent || '')
        .toLowerCase();

    console.log(txt);

    if(txt.includes('metal') || txt.includes('coil')){

        target = opt;
        break;
    }
}

if(!target){
    return "MATCHING OPTION NOT FOUND";
}

target.scrollIntoView({block:'center'});

await sleep(1000);

// ================= CLICK =================

target.dispatchEvent(new MouseEvent('mouseover',{
    bubbles:true,
    composed:true
}));

target.dispatchEvent(new MouseEvent('mousedown',{
    bubbles:true,
    composed:true
}));

target.dispatchEvent(new MouseEvent('mouseup',{
    bubbles:true,
    composed:true
}));

target.dispatchEvent(new MouseEvent('click',{
    bubbles:true,
    composed:true
}));

target.click();

return "OPTION CLICKED SUCCESSFULLY";
""")

time.sleep(5)

container = driver.execute_script("""

const containerText = arguments[0];

async function sleep(ms){
    return new Promise(r => setTimeout(r, ms));
}

async function run(){

    // =========================
    // DEEP QUERY
    // =========================

    function deepAll(selector, root=document){

        let results = [];

        try{
            results.push(...root.querySelectorAll(selector));
        }catch(e){}

        let all = [];

        try{
            all = root.querySelectorAll("*");
        }catch(e){}

        for(let el of all){

            if(el.shadowRoot){

                results.push(...deepAll(selector, el.shadowRoot));
            }
        }

        return results;
    }

    // =========================
    // INPUT
    // =========================

    let input = deepAll('input')
        .find(el =>
            (
                el.placeholder || ''
            )
            .toLowerCase()
            .includes('select container type')
        );

    if(!input){
        return "❌ INPUT NOT FOUND";
    }

    input.scrollIntoView({
        block:'center'
    });

    await sleep(1000);

    input.focus();
    input.click();

    await sleep(2000);

    // =========================
    // CLEAR
    // =========================

    input.value = "";

    input.dispatchEvent(new InputEvent('input',{
        bubbles:true,
        composed:true,
        data:"",
        inputType:'deleteContentBackward'
    }));

    await sleep(500);

    // =========================
    // TYPE VALUE
    // =========================

    for(let ch of containerText){

        input.value += ch;

        input.dispatchEvent(new InputEvent('input',{
            bubbles:true,
            composed:true,
            data: ch,
            inputType:'insertText'
        }));

        await sleep(150);
    }

    input.dispatchEvent(new Event('change',{
        bubbles:true,
        composed:true
    }));

    await sleep(3000);

    // =========================
    // FIND OPTION
    // =========================

    let options = deepAll('*');

    let target = null;

    for(let el of options){

        let txt = (
            el.innerText ||
            el.textContent ||
            ''
        )
        .replace(/\\s+/g,' ')
        .trim()
        .toLowerCase();

        if(
            txt === containerText.toLowerCase()
        ){
            target = el;
            break;
        }
    }

    if(!target){

        return "❌ OPTION NOT FOUND";
    }

    target.scrollIntoView({
        block:'center'
    });

    await sleep(1000);

    // =========================
    // CLICK OPTION
    // =========================

    if(target.shadowRoot){

        let btn = target.shadowRoot.querySelector("button");

        if(btn){

            btn.click();

        }else{

            target.click();
        }

    }else{

        target.click();
    }

    await sleep(2000);

    return {
        status : "✅ CONTAINER SELECTED",
        selected : containerText
    };
}

return run();

""", container_type)

time.sleep(5)

weight = driver.execute_script("""
async function sleep(ms){
    return new Promise(r => setTimeout(r, ms));
}

async function run(){

    // =========================
    // DEEP SEARCH
    // =========================

    function deepAll(selector, root=document){

        let results = [];

        try{
            results.push(...root.querySelectorAll(selector));
        }catch(e){}

        let all = [];

        try{
            all = root.querySelectorAll("*");
        }catch(e){}

        for(let el of all){

            if(el.shadowRoot){

                results.push(...deepAll(selector, el.shadowRoot));
            }
        }

        return results;
    }

    // =========================
    // FIND WEIGHT INPUT
    // =========================

    let input = deepAll('input')
        .find(el =>
            el.id === 'mc-input-weight'
        );

    if(!input){
        return "❌ WEIGHT INPUT NOT FOUND";
    }

    input.scrollIntoView({
        block:'center'
    });

    await sleep(1000);

    // =========================
    // CLICK INPUT FIRST
    // =========================

    input.focus();

    input.dispatchEvent(new MouseEvent('mousedown',{
        bubbles:true,
        composed:true
    }));

    input.dispatchEvent(new MouseEvent('mouseup',{
        bubbles:true,
        composed:true
    }));

    input.dispatchEvent(new MouseEvent('click',{
        bubbles:true,
        composed:true
    }));

    input.click();

    await sleep(1000);

    // =========================
    // CLEAR VALUE
    // =========================

    input.value = "";

    // =========================
    // TYPE 20000
    // =========================

    let text = "20000";

    for(let ch of text){

        input.value += ch;

        input.dispatchEvent(new InputEvent('input',{
            bubbles:true,
            composed:true,
            data: ch,
            inputType:'insertText'
        }));

        await sleep(120);
    }

    // trigger events
    input.dispatchEvent(new Event('change',{
        bubbles:true,
        composed:true
    }));

    input.dispatchEvent(new Event('blur',{
        bubbles:true,
        composed:true
    }));

    await sleep(1000);

    return "✅ 20000 ENTERED SUCCESSFULLY";
}

return run();
""")

time.sleep(5)

result = driver.execute_script("""

function deepAll(selector, root=document){

    let out = [];

    try{
        out.push(...root.querySelectorAll(selector));
    }catch(e){}

    let all = [];

    try{
        all = root.querySelectorAll("*");
    }catch(e){}

    for(let el of all){

        if(el.shadowRoot){
            out.push(...deepAll(selector, el.shadowRoot));
        }
    }

    return out;
}

// ====================================
// FIND mc-radio
// ====================================

let radio = deepAll('mc-radio[label="I am the price owner"]')[0];

if(!radio){
    return "❌ radio not found";
}

// ====================================
// GET SHADOW INPUT
// ====================================

let shadow = radio.shadowRoot;

if(!shadow){
    return "❌ shadowRoot missing";
}

let input = shadow.querySelector('input[type="radio"]');

if(!input){
    return "❌ input not found";
}

// ====================================
// SCROLL
// ====================================

input.scrollIntoView({
    block:'center'
});

// ====================================
// REAL CLICK
// ====================================

input.click();

// ====================================
// FORCE CHECKED
// ====================================

input.checked = true;

input.dispatchEvent(new Event('input',{
    bubbles:true,
    composed:true
}));

input.dispatchEvent(new Event('change',{
    bubbles:true,
    composed:true
}));

return "✅ clicked";

""")

time.sleep(10)
#***************************************************************************

retry = driver.execute_script("""
async function sleep(ms){
    return new Promise(r => setTimeout(r, ms));
}

function deepAll(selector, root=document){

    let results = [];

    try{
        results.push(...root.querySelectorAll(selector));
    }catch(e){}

    let all = [];

    try{
        all = root.querySelectorAll("*");
    }catch(e){}

    for(let el of all){

        if(el.shadowRoot){

            results.push(...deepAll(selector, el.shadowRoot));
        }
    }

    return results;
}

async function run(){

    let retryBtn = null;

    let buttons = deepAll('mc-button');

    for(let btn of buttons){

        let txt =
            (btn.getAttribute('label') || '') +
            ' ' +
            (btn.innerText || btn.textContent || '');

        txt = txt.toLowerCase();

        if(txt.includes('retry')){

            retryBtn = btn;
            break;
        }
    }

    // =========================================
    // BUTTON FOUND
    // =========================================

    if(retryBtn){

        retryBtn.scrollIntoView({
            behavior:'smooth',
            block:'center'
        });

        await sleep(1500);

        try{

            retryBtn.click();

        }catch(e){

            retryBtn.dispatchEvent(new MouseEvent('click',{
                bubbles:true,
                composed:true
            }));
        }

        await sleep(4000);

        return "RETRY_CLICKED";
    }

    // =========================================
    // BUTTON NOT FOUND
    // =========================================

    return "RETRY_NOT_FOUND";
}

return run();
""")

print(retry)

time.sleep(15)

spot = driver.execute_script("""
function deepAll(selector, root=document){

    let results = [];

    try{
        results.push(...root.querySelectorAll(selector));
    }catch(e){}

    let all = [];

    try{
        all = root.querySelectorAll("*");
    }catch(e){}

    for(let el of all){

        if(el.shadowRoot){
            results.push(...deepAll(selector, el.shadowRoot));
        }
    }

    return results;
}

let headers = deepAll('[data-test="pricing-info-header"]');

for(let h of headers){

    let txt = (h.innerText || h.textContent || '').trim();

    if(txt.includes('Spot and market rates')){

        return true;
    }
}

return false;
""")

# =========================================================
print(spot)
# =========================================================

if spot:

    print("✅ Spot and market rates FOUND")

    checkedCards = set()
    currentPage = 1

    while True:

        print(f"\n📄 PAGE : {currentPage}")
        print(f"✅ Checked Cards : {checkedCards}")

        rates = driver.execute_script("""

        async function sleep(ms){
            return new Promise(r => setTimeout(r, ms));
        }

        async function run(checkedCards){

            // =====================================================
            // DEEP QUERY
            // =====================================================

            function deepAll(selector, root=document){

                let results = [];

                try{
                    results.push(...root.querySelectorAll(selector));
                }catch(e){}

                let all = [];

                try{
                    all = root.querySelectorAll("*");
                }catch(e){}

                for(let el of all){

                    if(el.shadowRoot){

                        results.push(...deepAll(selector, el.shadowRoot));
                    }
                }

                return results;
            }

            // =====================================================
            // OFFER CARDS
            // =====================================================

            let cards = deepAll("div.offer-cards__body");

            if(cards.length === 0){

                return "NO_CARDS_FOUND";
            }

            // =====================================================
            // LOOP CARDS
            // =====================================================

            for(let i=0; i<cards.length; i++){

                if(checkedCards.includes(i)){
                    continue;
                }

                let card = cards[i];

                // =================================================
                // PRICE DETAILS BUTTON
                // =================================================

                let btn = null;

                let buttons = card.querySelectorAll("mc-button");

                for(let b of buttons){

                    let txt = (
                        (b.getAttribute("label") || "") + " " +
                        (b.textContent || "")
                    ).toLowerCase();

                    if(txt.includes("price details")){

                        btn = b;
                        break;
                    }
                }

                if(!btn){

                    checkedCards.push(i);
                    continue;
                }

                // =================================================
                // CLICK BUTTON
                // =================================================

                btn.scrollIntoView({
                    behavior: "smooth",
                    block: "center"
                });

                await sleep(2000);

                try{
                    btn.click();
                }
                catch(e){

                    btn.dispatchEvent(
                        new MouseEvent("click",{
                            bubbles:true,
                            composed:true
                        })
                    );
                }

                await sleep(8000);

                // =================================================
                // TABLE
                // =================================================

                let table =
                    deepAll(".mds-table table")[0];

                if(!table){

                    checkedCards.push(i);
                    continue;
                }

                // =================================================
                // DEPARTURE / ARRIVAL
                // =================================================

                let departure = "";
                let arrival = "";

                let modal =
                    deepAll(".offer-modal-header")[0];

                if(modal){

                    let dep =
                        modal.querySelector(
                            '[data-test="header-label-departure"]'
                        );

                    if(dep){

                        let parent = dep.parentElement;

                        let ps = parent.querySelectorAll("p");

                        if(ps.length > 1){

                            departure =
                                ps[1].textContent.trim();
                        }
                    }

                    let arr =
                        modal.querySelector(
                            '[data-test="header-label-arrival"]'
                        );

                    if(arr){

                        let parent = arr.parentElement;

                        let ps = parent.querySelectorAll("p");

                        if(ps.length > 1){

                            arrival =
                                ps[1].textContent.trim();
                        }
                    }
                }

                // =================================================
                // ROWS
                // =================================================

                let rows = table.querySelectorAll("tbody tr");

                console.log("TOTAL ROWS =", rows.length);

                let freightTotal = 0;
                let originTotal = 0;
                let destinationTotal = 0;

                let freightItems = [];
                let originItems = [];
                let destinationItems = [];

                let currentSection = "freight";

                // =================================================
                // LOOP ROWS
                // =================================================

                for(let row of rows){

                    let cols = row.querySelectorAll("td");

                    if(cols.length < 6){
                        continue;
                    }

                    // =============================================
                    // NAME
                    // =============================================

                    let chargeName =
                        cols[0].textContent
                        .replace(/\\s+/g,' ')
                        .trim();

                    if(!chargeName){
                        continue;
                    }

                    let lower =
                        chargeName.toLowerCase();

                    console.log("ROW =", chargeName);

                    // =============================================
                    // SECTION HEADINGS
                    // =============================================

                    if(lower === "origin charges"){

                        currentSection = "origin";
                        console.log("SECTION => ORIGIN");
                        continue;
                    }

                    if(lower === "destination charges"){

                        currentSection = "destination";
                        console.log("SECTION => DESTINATION");
                        continue;
                    }

                    if(lower === "freight charges"){

                        currentSection = "freight";
                        console.log("SECTION => FREIGHT");
                        continue;
                    }

                    // =============================================
                    // SKIP HEADER ROWS
                    // =============================================

                    if(
                        lower === "basis" ||
                        lower === "quantity" ||
                        lower === "currency" ||
                        lower === "unit price" ||
                        lower === "total price"
                    ){
                        continue;
                    }

                    // =============================================
                    // LAST COLUMN
                    // =============================================

                    let totalCol =
                        cols[cols.length - 1];

                    let rawText =
                        totalCol.textContent || "";

                    let amountText =
                        rawText
                        .replace(/,/g,'')
                        .replace(/[^0-9.]/g,'');

                    let amount =
                        parseFloat(amountText);

                    if(isNaN(amount)){
                        amount = 0;
                    }

                    console.log(
                        "CHARGE =",
                        chargeName,
                        "SECTION =",
                        currentSection,
                        "AMOUNT =",
                        amount
                    );

                    let item = {
                        charge: chargeName,
                        amount: amount
                    };

                    // =============================================
                    // SAVE DATA
                    // =============================================

                    if(currentSection === "freight"){

                        freightTotal += amount;
                        freightItems.push(item);
                    }

                    else if(currentSection === "origin"){

                        originTotal += amount;
                        originItems.push(item);
                    }

                    else if(currentSection === "destination"){

                        destinationTotal += amount;
                        destinationItems.push(item);
                    }
                }

                // =================================================
                // CLOSE MODAL
                // =================================================

                let closeBtn =
                    deepAll(
                        'mc-button[data-test="offer-modal-close-icon"]'
                    )[0];

                if(closeBtn){

                    try{
                        closeBtn.click();
                    }
                    catch(e){}

                    await sleep(3000);
                }

                checkedCards.push(i);

                // =================================================
                // RETURN
                // =================================================

                return {

                    status: "SUCCESS",

                    departure,
                    arrival,

                    freightTotal,
                    originTotal,
                    destinationTotal,

                    freightItems,
                    originItems,
                    destinationItems,

                    checkedCards
                };
            }

            // =====================================================
            // ALL CARDS CHECKED
            // =====================================================

            return {
                status: "ALL_CARDS_CHECKED",
                checkedCards
            };
        }

        return run(arguments[0]);

        """, list(checkedCards))

        print(rates)

        # =========================================================
        # SUCCESS
        # =========================================================

        if isinstance(rates, dict):

            status = rates.get("status")

            # =====================================================
            # SUCCESS
            # =====================================================

            if status == "SUCCESS":

                checkedCards = set(
                    rates.get("checkedCards", [])
                )

                departure = rates.get("departure")
                arrival = rates.get("arrival")

                FREIGHT = round(rates.get("freightTotal", 0), 1)
                ORIGIN = round(rates.get("originTotal", 0), 1)
                DESTINATION = round(rates.get("destinationTotal", 0) / 96, 1)

                freightItems = rates.get("freightItems", [])
                originItems = rates.get("originItems", [])
                destinationItems = rates.get("destinationItems", [])

                export = 0

                total = round(
                    FREIGHT + ORIGIN + export + DESTINATION,
                    1
                )

                print("\n========================")
                print("✅ FINAL DATA")
                print("========================")

                print("Departure :", departure)
                print("Arrival :", arrival)

                # =================================================
                # FREIGHT
                # =================================================

                print("\n========================")
                print("🚢 FREIGHT CHARGES")
                print("========================")

                for item in freightItems:

                    print(
                        f"{item['charge']} : {item['amount']}"
                    )

                print("TOTAL FREIGHT :", FREIGHT)

                # =================================================
                # ORIGIN
                # =================================================

                print("\n========================")
                print("🏭 ORIGIN CHARGES")
                print("========================")

                for item in originItems:

                    print(
                        f"{item['charge']} : {item['amount']}"
                    )

                print("TOTAL ORIGIN :", ORIGIN)

                # =================================================
                # DESTINATION
                # =================================================

                print("\n========================")
                print("🏁 DESTINATION CHARGES")
                print("========================")

                for item in destinationItems:

                    print(
                        f"{item['charge']} : {item['amount']}"
                    )

                print("TOTAL DESTINATION :", DESTINATION)

                # =================================================
                # GRAND TOTAL
                # =================================================

                print("\n========================")
                print("💰 GRAND TOTAL")
                print("========================")

                print("TOTAL :", total)

                # =================================================
                # SAVE
                # =================================================

                data.append([

                    today,
                    "MAERSK",

                    origin_port_name,
                    destination_port_name,

                    cont_size,
                    20000,

                    departure,
                    arrival,

                    FREIGHT,
                    ORIGIN,
                    export,
                    DESTINATION,

                    total,                

                    ""
                ])

            # =====================================================
            # NEXT PAGE
            # =====================================================

            elif status == "ALL_CARDS_CHECKED":

                print("➡️ ALL PRICE DETAILS CHECKED")

                next_clicked = driver.execute_script("""

                async function sleep(ms){
                    return new Promise(r => setTimeout(r, ms));
                }

                async function run(){

                    function deepAll(selector, root=document){

                        let results = [];

                        try{
                            results.push(...root.querySelectorAll(selector));
                        }catch(e){}

                        let all = [];

                        try{
                            all = root.querySelectorAll("*");
                        }catch(e){}

                        for(let el of all){

                            if(el.shadowRoot){

                                results.push(
                                    ...deepAll(selector, el.shadowRoot)
                                );
                            }
                        }

                        return results;
                    }

                    // =============================================
                    // FIND NEXT BUTTON
                    // =============================================

                    let nextBtn = null;

                    let buttons = deepAll("mc-button");

                    for(let b of buttons){

                        let txt = (
                            (b.getAttribute("label") || "") + " " +
                            (b.textContent || "")
                        ).toLowerCase();

                        if(txt.includes("next")){

                            nextBtn = b;
                            break;
                        }
                    }

                    if(!nextBtn){

                        return "NO_NEXT_BUTTON";
                    }

                    // =============================================
                    // DISABLED CHECK
                    // =============================================

                    let disabled = false;

                    try{

                        if(
                            nextBtn.hasAttribute("disabled")
                        ){
                            disabled = true;
                        }

                        if(nextBtn.shadowRoot){

                            let realBtn =
                                nextBtn.shadowRoot.querySelector("button");

                            if(realBtn && realBtn.disabled){

                                disabled = true;
                            }
                        }

                    }catch(e){}

                    if(disabled){

                        return "NO_MORE_PAGES";
                    }

                    // =============================================
                    // CLICK NEXT
                    // =============================================

                    nextBtn.scrollIntoView({
                        behavior: "smooth",
                        block: "center"
                    });

                    await sleep(2000);

                    try{
                        nextBtn.click();
                    }
                    catch(e){

                        nextBtn.dispatchEvent(
                            new MouseEvent("click",{
                                bubbles:true,
                                composed:true
                            })
                        );
                    }

                    await sleep(8000);

                    return "NEXT_PAGE_CLICKED";
                }

                return run();

                """)

                print(next_clicked)

                if next_clicked == "NEXT_PAGE_CLICKED":

                    currentPage += 1

                    checkedCards = set()

                    time.sleep(5)

                    continue

                else:

                    print("❌ ALL PAGES COMPLETED")
                    break

        # =========================================================
        # NO CARDS
        # =========================================================

        elif rates == "NO_CARDS_FOUND":

            print("❌ NO OFFER CARD FOUND")
            break

else:

    print("❌ Spot and market rates NOT FOUND")

    # **********************************************************
    # SELECT TOMORROW
    # **********************************************************

    select_tomorrow_result = driver.execute_script("""
    async function sleep(ms){
        return new Promise(r => setTimeout(r, ms));
    }

    async function run(){

        function deepAll(selector, root=document){

            let results = [];

            try{
                results.push(...root.querySelectorAll(selector));
            }catch(e){}

            let all = [];

            try{
                all = root.querySelectorAll("*");
            }catch(e){}

            for(let el of all){

                if(el.shadowRoot){

                    results.push(...deepAll(selector, el.shadowRoot));
                }
            }

            return results;
        }

        let selectTomorrowBtn = null;

        let links = deepAll('a');

        for(let link of links){

            let text = (link.innerText || '').trim();

            let cls = link.className || '';

            if(
                text.includes('Select tomorrow') ||
                cls.includes('select-tomorrow-link')
            ){

                selectTomorrowBtn = link;
                break;
            }
        }

        if(!selectTomorrowBtn){
            return "❌ SELECT TOMORROW BUTTON NOT FOUND";
        }

        selectTomorrowBtn.scrollIntoView({
            block:'center'
        });

        await sleep(1000);

        selectTomorrowBtn.dispatchEvent(new MouseEvent('mouseover',{
            bubbles:true,
            composed:true
        }));

        selectTomorrowBtn.dispatchEvent(new MouseEvent('mousedown',{
            bubbles:true,
            composed:true
        }));

        selectTomorrowBtn.dispatchEvent(new MouseEvent('mouseup',{
            bubbles:true,
            composed:true
        }));

        selectTomorrowBtn.dispatchEvent(new MouseEvent('click',{
            bubbles:true,
            composed:true
        }));

        selectTomorrowBtn.click();

        await sleep(3000);

        return "✅ SELECT TOMORROW CLICKED";
    }

    return run();
    """)

    print(select_tomorrow_result)

    # ==========================================================
    # IF SELECT TOMORROW FAILED
    # ==========================================================

    if "NOT FOUND" in str(select_tomorrow_result):

        print("❌ SELECT TOMORROW FAILED")

        data.append([
            today,
            "MAERSK",
            origin_port_name,
            destination_port_name,
            cont_size,
            20000,
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "No Data"
        ])

    # ==========================================================
    # CONTINUE FLOW
    # ==========================================================

    else:

        time.sleep(10)

        # **********************************************************
        # CONTINUE BUTTON
        # **********************************************************

        continue_result = driver.execute_script("""

        async function sleep(ms){
            return new Promise(r => setTimeout(r, ms));
        }

        async function run(){

            function deepQuery(selector){

                function search(root){

                    let el = root.querySelector(selector);

                    if(el) return el;

                    let all = root.querySelectorAll("*");

                    for(let node of all){

                        if(node.shadowRoot){

                            let found = search(node.shadowRoot);

                            if(found) return found;
                        }
                    }

                    return null;
                }

                return search(document);
            }

            let mcBtn = deepQuery("#od3cpContinueButton");

            if(!mcBtn)
                return "❌ MC BUTTON NOT FOUND";

            let realBtn = null;

            if(mcBtn.shadowRoot){
                realBtn = mcBtn.shadowRoot.querySelector("button");
            }

            if(!realBtn)
                return "❌ INNER BUTTON NOT FOUND";

            mcBtn.disabled = false;
            realBtn.disabled = false;

            mcBtn.removeAttribute("disabled");
            realBtn.removeAttribute("disabled");

            realBtn.style.pointerEvents = "auto";
            realBtn.style.opacity = "1";
            realBtn.style.visibility = "visible";

            realBtn.scrollIntoView({
                behavior: "smooth",
                block: "center"
            });

            await sleep(1500);

            let events = [
                "mouseover",
                "mouseenter",
                "pointerdown",
                "mousedown",
                "focus",
                "pointerup",
                "mouseup",
                "click"
            ];

            for(let ev of events){

                realBtn.dispatchEvent(
                    new MouseEvent(ev,{
                        bubbles:true,
                        cancelable:true,
                        composed:true,
                        view:window
                    })
                );

                await sleep(100);
            }

            realBtn.click();

            await sleep(3000);

            return {
                status : "✅ CONTINUE BUTTON CLICKED",
                currentUrl : location.href,
                buttonText : realBtn.innerText,
                disabled : realBtn.disabled
            };
        }

        return run();
        """)

        print(continue_result)

        time.sleep(10)

        # **********************************************************
        # LOAD MORE SAILINGS
        # **********************************************************

        driver.set_script_timeout(300)

        load_more_result = driver.execute_async_script("""

        const callback = arguments[arguments.length - 1];

        async function sleep(ms){
            return new Promise(r => setTimeout(r, ms));
        }

        async function run(){

            function deepAll(selector, root=document){

                let results = [];

                try{
                    results.push(...root.querySelectorAll(selector));
                }catch(e){}

                let els = [];

                try{
                    els = root.querySelectorAll("*");
                }catch(e){}

                for(let el of els){

                    if(el.shadowRoot){
                        results.push(...deepAll(selector, el.shadowRoot));
                    }
                }

                return results;
            }

            let total = 0;

            while(true){

                let allBtns = deepAll("mc-button");

                let btn = null;

                for(let b of allBtns){

                    let txt = (
                        b.innerText ||
                        b.textContent ||
                        b.getAttribute("label") ||
                        ""
                    ).trim();

                    if(txt.includes("Search more sailing options")){

                        btn = b;
                        break;
                    }
                }

                if(!btn){

                    callback("DONE | TOTAL CLICKS = " + total);
                    return;
                }

                try{

                    let realBtn =
                        btn.shadowRoot.querySelector("button");

                    realBtn.scrollIntoView({
                        behavior:"smooth",
                        block:"center"
                    });

                    await sleep(2000);

                    realBtn.click();

                    total++;

                    console.log("CLICKED =>", total);

                    await sleep(8000);

                }catch(e){

                    callback("ERROR => " + e);
                    return;
                }
            }
        }

        run();
        """)

        print(load_more_result)

        # **********************************************************
        # TOTAL CARDS
        # **********************************************************

        sailing_cards_length = driver.execute_script("""

        return document.querySelectorAll(
            'article.new-sailings-card-article'
        ).length;

        """)

        print(
            "TOTAL SAILING CARDS :",
            sailing_cards_length
        )

        # ==========================================================
        # NO CARDS
        # ==========================================================

        if sailing_cards_length == 0:

            data.append([
                today,
                "MAERSK",
                origin_port_name,
                destination_port_name,
                cont_size,
                20000,
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "No Data"
            ])

        # ==========================================================
        # PROCESS ALL CARDS
        # ==========================================================

        else:

            for card_index in range(sailing_cards_length):

                print(
                    f"\n{'='*20} "
                    f"CARD {card_index + 1} "
                    f"{'='*20}"
                )

                # **************************************************
                # YOUR EXISTING CARD LOGIC
                # **************************************************

                sailing_data = driver.execute_script(f"""
                // YOUR EXISTING JS
                """)

                print(sailing_data)

                if not sailing_data:

                    print("❌ No card data found")
                    continue

                # **************************************************
                # DATE FORMAT
                # **************************************************

                dep_raw = (
                    sailing_data["departure"]
                    .split(",")[0]
                    .strip()
                )

                arr_raw = (
                    sailing_data["arrival"]
                    .split(",")[0]
                    .strip()
                )

                try:

                    dep_date = datetime.strptime(
                        dep_raw,
                        "%d %b %Y"
                    )

                    arr_date = datetime.strptime(
                        arr_raw,
                        "%d %b %Y"
                    )

                    start_date = dep_date.strftime(
                        "%d-%b-%Y"
                    )

                    end_date = arr_date.strftime(
                        "%d-%b-%Y"
                    )

                except Exception as e:

                    print("❌ Date parse error :", e)

                    start_date = ""
                    end_date = ""

                # **************************************************
                # PRICE
                # **************************************************

                price_text = (
                    sailing_data["price"].strip()
                )

                parts = price_text.split()

                currency = (
                    parts[0]
                    if len(parts) > 0
                    else ""
                )

                amount = (
                    int(float(parts[1]))
                    if len(parts) > 1
                    else 0
                )

                print("AMOUNT :", amount)

                if amount <= 0:

                    print(
                        "❌ Amount not found"
                    )

                    continue

                # **************************************************
                # PRICE BREAKDOWN CLICK
                # **************************************************

                click_result = driver.execute_async_script(f"""
                // YOUR EXISTING JS
                """)

                print(
                    "\nPRICE BREAKDOWN CLICK :",
                    click_result
                )

                time.sleep(2)

                # **************************************************
                # PRICE BREAKDOWN DATA
                # **************************************************

                price_breakdown_data = driver.execute_script("""
                // YOUR EXISTING JS
                """)

                print(price_breakdown_data)

                # **************************************************
                # TOTALS
                # **************************************************

                if "freight_total" in price_breakdown_data:

                    FREIGHT = round(
                        price_breakdown_data[
                            "freight_total"
                        ],
                        1
                    )

                    ORIGIN = round(
                        price_breakdown_data[
                            "origin_total"
                        ],
                        1
                    )

                    DESTINATION = round(
                        price_breakdown_data[
                            "destination_total"
                        ] / 96,
                        1
                    )

                    export = 0

                    total = round(
                        FREIGHT +
                        ORIGIN +
                        DESTINATION,
                        1
                    )

                else:

                    FREIGHT = 0
                    ORIGIN = 0
                    DESTINATION = 0
                    export = 0
                    total = 0

                # **************************************************
                # SAVE DATA
                # **************************************************

                data.append([

                    today,
                    "MAERSK",

                    origin_port_name,
                    destination_port_name,

                    cont_size,
                    20000,

                    start_date,
                    end_date,

                    FREIGHT,
                    ORIGIN,
                    export,
                    DESTINATION,

                    total,

                    ""
                ])

# =========================================================
# SAVE EXCEL
# =========================================================

new_df = pd.DataFrame(data, columns=columns)

try:
    old_df = pd.read_excel(file_path)
    final_df = pd.concat([old_df, new_df], ignore_index=True)

except:
    final_df = new_df

final_df.to_excel(file_path, index=False)

print("✅ Data saved successfully")

driver.quit()