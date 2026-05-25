import os
import pandas as pd
from django.shortcuts import render, redirect
from .models import ShippingRate
from django.core.paginator import Paginator
from datetime import datetime, date
import subprocess
from django.http import JsonResponse
from django.contrib import messages

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
rate_file_path = os.path.join(BASE_DIR, "global", "rates.xlsx")

def import_excel(request):    
   
    df = pd.read_excel(rate_file_path)

    df = df.fillna("")
    df = df.drop_duplicates()

    ShippingRate.objects.all().delete()

    def clean(val):
        if pd.isna(val):
            return ""
        return str(val).strip()

    def parse_date(val):
        try:
            return pd.to_datetime(val).date()
        except:
            return None

    objects = []

    for _, row in df.iterrows():
        objects.append(ShippingRate(

            # ✅ ALL COLUMN MAPPING
            date = parse_date(row.get('date')),

            shipping_line = clean(row.get('ship_line')),
            origin_port_name = clean(row.get('org_port')),
            destination_port_name = clean(row.get('dest_port')),

            container_size = clean(row.get('cont_type')),
            kgs = int(row.get('weight')) if str(row.get('weight')).strip() not in ["", "nan"] else None,

            valid_from_etd = clean(row.get('etd_from')),
            valid_to_eta = clean(row.get('eta_to')),

            ocean_freight = clean(row.get('ocean_frt')),
            freight_surcharge = clean(row.get('frt_surchg')),
            export_surcharge = clean(row.get('exp_surchg')),
            import_surcharge = clean(row.get('imp_surchg')),

            total_cost_usd = (
                float(str(row.get('total_usd')).replace(",", "").strip())
                if str(row.get('total_usd')).strip() not in ["", "nan"]
                else None
            ),

            remarks = clean(row.get('remarks')),
        ))

    ShippingRate.objects.bulk_create(objects)

    messages.success(request, "Excel Imported Successfully")
    return redirect('rates')


def success_page(request):
    return render(request, 'success.html')    

# =========================
# COMMON FILTER FUNCTION
# =========================
def get_filtered_queryset(request):
    rates_list = ShippingRate.objects.all().order_by('-date')

    selected_filter = request.GET.get('filter')
    date_filter = request.GET.get('date')
    shipping_line = request.GET.get('shipping_line')
    origin = request.GET.get('origin')
    destination = request.GET.get('destination')
    container = request.GET.get('container')

    # ✅ DATE FILTER
    if selected_filter == "today":
        rates_list = rates_list.filter(date=date.today())

    elif date_filter:
        try:
            parsed_date = datetime.strptime(date_filter, "%Y-%m-%d").date()
            rates_list = rates_list.filter(date=parsed_date)
        except:
            pass

    # ✅ OTHER FILTERS
    if shipping_line:
        rates_list = rates_list.filter(shipping_line=shipping_line)

    if origin:
        rates_list = rates_list.filter(origin_port_name__iexact=origin.strip())

    if destination:
        rates_list = rates_list.filter(destination_port_name__iexact=destination.strip())

    if container:
        rates_list = rates_list.filter(container_size=container)

    return rates_list

# =========================
# CHEAPEST LOGIC (REUSABLE)
# =========================

def apply_cheapest_filter(rates_list):
    cheapest_map = {}

    for r in rates_list:
        if not r.total_cost_usd:
            continue

        key = (
            r.shipping_line,
            r.origin_port_name,
            r.destination_port_name,
            r.container_size
        )

        price = float(r.total_cost_usd)

        if key not in cheapest_map or price < cheapest_map[key][1]:
            cheapest_map[key] = (r.id, price)

    cheapest_ids = [v[0] for v in cheapest_map.values()]

    return rates_list.filter(id__in=cheapest_ids), cheapest_ids


# =========================
# MAIN VIEW
# =========================
def show_rates(request):

    # 🔹 GET FILTER VALUES (MULTI SELECT)
    selected_filter = request.GET.get('filter')
    date_filter = request.GET.get('date')

    selected_shipping = request.GET.getlist('shipping_line')
    selected_origin = request.GET.getlist('origin')
    selected_destination = request.GET.getlist('destination')
    selected_container = request.GET.getlist('container')

    # 🔹 BASE QUERY
    rates_list = ShippingRate.objects.all().order_by('-date')

    # =========================
    # DATE FILTER
    # =========================
    if selected_filter == "today":
        rates_list = rates_list.filter(date=date.today())

    elif date_filter:
        try:
            parsed_date = datetime.strptime(date_filter, "%Y-%m-%d").date()
            rates_list = rates_list.filter(date=parsed_date)
        except:
            pass

    # =========================
    # MULTI FILTER (IMPORTANT FIX)
    # =========================
    if selected_shipping:
        rates_list = rates_list.filter(shipping_line__in=selected_shipping)

    if selected_origin:
        rates_list = rates_list.filter(origin_port_name__in=selected_origin)

    if selected_destination:
        rates_list = rates_list.filter(destination_port_name__in=selected_destination)

    if selected_container:
        rates_list = rates_list.filter(container_size__in=selected_container)

    # =========================
    # CHEAPEST LOGIC
    # =========================
    cheapest_ids = []

    if selected_filter == "cheapest":
        rates_list, cheapest_ids = apply_cheapest_filter(rates_list)

    else:
        temp_map = {}

        for r in rates_list:
            if not r.total_cost_usd:
                continue

            key = (
                r.shipping_line,
                r.origin_port_name,
                r.destination_port_name,
                r.container_size
            )

            price = float(r.total_cost_usd)

            if key not in temp_map or price < temp_map[key][1]:
                temp_map[key] = (r.id, price)

        cheapest_ids = [v[0] for v in temp_map.values()]

    # =========================
    # PAGINATION
    # =========================
    paginator = Paginator(rates_list, 20)
    page_number = request.GET.get('page')
    rates = paginator.get_page(page_number)

    # =========================
    # CLEAN DATA FORMAT
    # =========================
    for r in rates:
        r.valid_from_etd = str(r.valid_from_etd).strip() if r.valid_from_etd else "-"
        r.valid_to_eta = str(r.valid_to_eta).strip() if r.valid_to_eta else "-"
        r.remarks = str(r.remarks).strip() if r.remarks else "-"

    # =========================
    # DROPDOWN DATA (DISTINCT)
    # =========================
    shipping_lines = ShippingRate.objects.values_list('shipping_line', flat=True).distinct()
    origins = ShippingRate.objects.values_list('origin_port_name', flat=True).distinct()
    destinations = ShippingRate.objects.values_list('destination_port_name', flat=True).distinct()
    containers = ShippingRate.objects.values_list('container_size', flat=True).distinct()

    # =========================
    # RENDER
    # =========================
    return render(request, 'rates.html', {
        'rates': rates,
        'cheapest_ids': cheapest_ids,

        # dropdown data
        'shipping_lines': shipping_lines,
        'origins': origins,
        'destinations': destinations,
        'containers': containers,

        # selected values (IMPORTANT for checkbox retain)
        'selected_shipping': selected_shipping,
        'selected_origin': selected_origin,
        'selected_destination': selected_destination,
        'selected_container': selected_container,
    })
# =========================
# DOWNLOAD EXCEL VIEW
# =========================
from django.http import HttpResponse

def download_excel(request):

    # =========================
    # GET FILTER VALUES (MULTI)
    # =========================
    selected_filter = request.GET.get('filter')
    date_filter = request.GET.get('date')

    selected_shipping = request.GET.getlist('shipping_line')
    selected_origin = request.GET.getlist('origin')
    selected_destination = request.GET.getlist('destination')
    selected_container = request.GET.getlist('container')

    # =========================
    # BASE QUERY
    # =========================
    queryset = ShippingRate.objects.all().order_by('-date')

    # =========================
    # DATE FILTER
    # =========================
    if selected_filter == "today":
        queryset = queryset.filter(date=date.today())

    elif date_filter:
        try:
            parsed_date = datetime.strptime(date_filter, "%Y-%m-%d").date()
            queryset = queryset.filter(date=parsed_date)
        except:
            pass

    # =========================
    # MULTI FILTER
    # =========================
    if selected_shipping:
        queryset = queryset.filter(shipping_line__in=selected_shipping)

    if selected_origin:
        queryset = queryset.filter(origin_port_name__in=selected_origin)

    if selected_destination:
        queryset = queryset.filter(destination_port_name__in=selected_destination)

    if selected_container:
        queryset = queryset.filter(container_size__in=selected_container)

    # =========================
    # CHEAPEST FILTER
    # =========================
    if selected_filter == "cheapest":
        queryset, _ = apply_cheapest_filter(queryset)

    # =========================
    # DATAFRAME
    # =========================
    data = queryset.values()
    df = pd.DataFrame(data)

    if df.empty:
        df = pd.DataFrame(columns=[
            'Date','Shipping','Origin','Destination','Container',
            'KGS','Ocean','FRT','EXP','IMP',
            'ETD','ETA','Total ($)','Remarks'
        ])

    # REMOVE ID
    if 'id' in df.columns:
        df.drop(columns=['id'], inplace=True)

    # RENAME
    df.rename(columns={
        'date': 'Date',
        'shipping_line': 'Shipping',
        'origin_port_name': 'Origin',
        'destination_port_name': 'Destination',
        'container_size': 'Container',
        'kgs': 'KGS',
        'ocean_freight': 'Ocean',
        'freight_surcharge': 'FRT',
        'export_surcharge': 'EXP',
        'import_surcharge': 'IMP',
        'valid_from_etd': 'ETD',
        'valid_to_eta': 'ETA',
        'total_cost_usd': 'Total ($)',
        'remarks': 'Remarks',
    }, inplace=True)

    # ORDER
    columns_order = [
        'Date','Shipping','Origin','Destination','Container',
        'KGS','Ocean','FRT','EXP','IMP',
        'ETD','ETA','Total ($)','Remarks'
    ]

    df = df[[col for col in columns_order if col in df.columns]]

    # =========================
    # FILE RESPONSE
    # =========================
    today = datetime.now().strftime("%d-%m-%Y")
    filename = f"SR_{today}.xlsx"

    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = f'attachment; filename="{filename}"'

    df.to_excel(response, index=False)

    return response
# =========================
# FILE PATHS
# =========================

file_path_shipping = os.path.join(BASE_DIR, "global", "shipping_details.xlsx")
SCRAPERS_DIR = os.path.join(BASE_DIR, "scrapers")

SCRAPERS = {
    "HAPAG": os.path.join(SCRAPERS_DIR, "hapaglloyd.py"),
    "CMA CGM": os.path.join(SCRAPERS_DIR, "cmacgm.py"),
    "ICONTAINERS": os.path.join(SCRAPERS_DIR, "icontainers.py"),
    "MSC": os.path.join(SCRAPERS_DIR, "msc.py"),
    "MAERSK": os.path.join(SCRAPERS_DIR, "maerks.py"),
}

COLUMN_MAP = {
    "HAPAG": "hapag",
    "CMA CGM": "cmacgm",
    "ICONTAINERS": "icontainers",
    "MSC": "msc",
    "MAERSK": "maersk"
}

# =========================
# SCRAPER PAGE
# =========================
def scraper_page(request):
    return render(request, "scraper.html", {
        "shipping_lines": list(COLUMN_MAP.keys())
    })


# =========================
# GET PORTS + SIZE
# =========================
def get_ports(request):
    containers = request.GET.getlist("container")
    shipping_lines = request.GET.getlist("shipping_line")

    try:
        port_df = pd.read_excel(file_path_shipping, sheet_name="port")
        size_df = pd.read_excel(file_path_shipping, sheet_name="size")

        # Clean columns
        port_df.columns = port_df.columns.str.strip().str.lower()
        size_df.columns = size_df.columns.str.strip().str.lower()

        # ✅ applymap हटाकर safe cleaning
        for col in port_df.columns:
            port_df[col] = port_df[col].astype(str).str.strip()

        for col in size_df.columns:
            size_df[col] = size_df[col].astype(str).str.strip()

        # =========================
        # CONTAINERS
        # =========================
        containers_list = (
            size_df["size"]
            .dropna()
            .astype(str)
            .str.strip()
            .drop_duplicates()
            .tolist()
        )

        # =========================
        # SIZE MAPPING
        # =========================
        size_mapping = {} 

        for _, row in size_df.iterrows():

            key = str(row["size"]).strip()

            size_mapping[key] = {}

            for col in COLUMN_MAP.values():

                val = row.get(col)

                 # ✅ skip empty / nan
                if pd.isna(val):
                    continue
                val = str(val).strip()
                if not val or val.lower() == "nan":
                    continue
                size_mapping[key][col] = val

        # =========================
        # FILTER SHIPPING LINES
        # =========================
        filtered_shipping_lines = list(COLUMN_MAP.keys())

        if containers:
            filtered_shipping_lines = [
                line for line, col in COLUMN_MAP.items()
                if any(col in size_mapping.get(c, {}) for c in containers)
            ]

        # =========================
        # PORTS LOAD
        # =========================
        ports_set = set()

        if shipping_lines:
            for line in shipping_lines:
                col_key = COLUMN_MAP.get(line)
                if not col_key:
                    continue

                col_name = col_key.lower()

                if col_name in port_df.columns:
                    df_filtered = port_df[[col_name, "port"]].dropna()
                    df_filtered = df_filtered.drop_duplicates(subset=["port"])
                    ports_set.update(df_filtered["port"].tolist())
        else:
            ports_set = set(port_df["port"].dropna().tolist())

        ports = sorted(list(ports_set))

        return JsonResponse({
            "containers": containers_list,
            "shipping_lines": filtered_shipping_lines,
            "origins": ports,
            "destinations": ports,
            "size_mapping": size_mapping
        })

    except Exception as e:
        return JsonResponse({"error": str(e)}, status=500)


# ================= IMPORT =================

import subprocess
import threading
import os
import time
import json
import pandas as pd

from django.http import JsonResponse
from django.views.decorators.csrf import csrf_exempt


# ================= GLOBAL =================

RUNNING_SCRAPERS = {}

PENDING_SCRAPERS = []

CURRENT_RUNNING = None

STOPPED_SCRAPERS = set()

# COMPLETE HISTORY
COMPLETED_SCRAPERS = []


# ================= RUN SCRAPER =================

@csrf_exempt
def run_scraper(request):

    global RUNNING_SCRAPERS
    global PENDING_SCRAPERS
    global CURRENT_RUNNING
    global STOPPED_SCRAPERS
    global COMPLETED_SCRAPERS

    if request.method != "POST":

        return JsonResponse({
            "status": "error",
            "msg": "Invalid Request"
        })

    shipping_lines = request.POST.getlist("shipping_line[]")
    origins = request.POST.getlist("origin[]")
    destinations = request.POST.getlist("destination[]")
    containers = request.POST.getlist("container[]")

    if not shipping_lines:

        return JsonResponse({
            "status": "error",
            "msg": "No Shipping Line Selected"
        })

    # ================= RESET =================

    STOPPED_SCRAPERS.clear()

    COMPLETED_SCRAPERS.clear()

    # ================= SAVE PENDING =================

    PENDING_SCRAPERS = shipping_lines.copy()

    def scraper_worker():

        global CURRENT_RUNNING
        global PENDING_SCRAPERS
        global STOPPED_SCRAPERS
        global COMPLETED_SCRAPERS

        try:

            # ================= READ EXCEL =================

            port_df = pd.read_excel(
                file_path_shipping,
                sheet_name="port"
            )

            size_df = pd.read_excel(
                file_path_shipping,
                sheet_name="size"
            )

            port_df.columns = (
                port_df.columns.str.strip().str.lower()
            )

            size_df.columns = (
                size_df.columns.str.strip().str.lower()
            )

            # ================= LOOP SHIPPING LINES =================

            for line in shipping_lines:

                try:

                    # ================= STOP CHECK =================

                    if line in STOPPED_SCRAPERS:

                        if line in PENDING_SCRAPERS:
                            PENDING_SCRAPERS.remove(line)

                        continue

                    # ================= CURRENT RUNNING =================

                    CURRENT_RUNNING = line

                    # REMOVE FROM PENDING
                    if line in PENDING_SCRAPERS:
                        PENDING_SCRAPERS.remove(line)

                    # ================= VALUES =================

                    origin_name = origins[0]

                    dest_name = destinations[0]

                    container = containers[0]

                    # ================= COLUMN =================

                    column = COLUMN_MAP.get(line)

                    if not column:

                        print(f"Column not found for {line}")

                        CURRENT_RUNNING = None

                        continue

                    column = column.lower()

                    # ================= PORT ROW =================

                    origin_row = port_df[
                        port_df["port"] == origin_name
                    ]

                    dest_row = port_df[
                        port_df["port"] == dest_name
                    ]

                    if origin_row.empty or dest_row.empty:

                        print(f"Port mapping not found for {line}")

                        CURRENT_RUNNING = None

                        continue

                    # ================= PORT CODE =================

                    origin_code = origin_row.iloc[0][column]

                    dest_code = dest_row.iloc[0][column]

                    # ================= CONTAINER MAPPING =================

                    container_row = size_df[
                        size_df["size"] == container
                    ]

                    if container_row.empty:

                        print(
                            f"Container mapping not found for {container}"
                        )

                        CURRENT_RUNNING = None

                        continue

                    final_container = container_row.iloc[0][column]

                    print(
                        f"{line} -> {container} => {final_container}"
                    )

                    # ================= SCRIPT PATH =================

                    script_path = SCRAPERS.get(line)

                    if (
                        not script_path
                        or
                        not os.path.exists(script_path)
                    ):

                        print(f"Script not found for {line}")

                        CURRENT_RUNNING = None

                        continue

                    # ================= START PROCESS =================

                    process = subprocess.Popen(

                        [
                            "python",
                            script_path,
                            str(origin_code),
                            str(origin_name),
                            str(dest_code),
                            str(dest_name),
                            str(final_container)
                        ]

                    )

                    RUNNING_SCRAPERS[line] = process

                    # ================= WAIT =================

                    try:

                        return_code = process.wait()

                        print(
                            f"{line} finished with code {return_code}"
                        )

                    finally:

                        # REMOVE RUNNING
                        if line in RUNNING_SCRAPERS:
                            del RUNNING_SCRAPERS[line]

                        # COMPLETED
                        if (
                            process.returncode == 0
                            and
                            line not in STOPPED_SCRAPERS
                        ):

                            COMPLETED_SCRAPERS.append(line)

                        CURRENT_RUNNING = None

                    # ================= NEXT DELAY =================

                    time.sleep(2)

                except Exception as err:

                    print(f"{line} Error: {err}")

                    CURRENT_RUNNING = None

                    continue

        except Exception as e:

            print(f"Worker Error: {e}")

            CURRENT_RUNNING = None

    # ================= THREAD START =================

    threading.Thread(
        target=scraper_worker,
        daemon=True
    ).start()

    return JsonResponse({

        "status": "success",

        "msg": "Scraper Started Successfully"

    })

# ================= GET RUNNING SCRAPERS =================

def get_running_scrapers(request):

    global RUNNING_SCRAPERS
    global CURRENT_RUNNING
    global PENDING_SCRAPERS
    global COMPLETED_SCRAPERS

    active = []

    remove_keys = []

    for line, process in RUNNING_SCRAPERS.items():

        if process.poll() is None:

            active.append(line)

        else:

            remove_keys.append(line)

    for k in remove_keys:

        del RUNNING_SCRAPERS[k]

    return JsonResponse({

        "running_scrapers": active,

        "current_running": CURRENT_RUNNING,

        "pending_scrapers": PENDING_SCRAPERS,

        "completed_scrapers": COMPLETED_SCRAPERS

    })


# ================= STOP SELECTED =================

@csrf_exempt
def stop_selected_scrapers(request):

    global RUNNING_SCRAPERS
    global CURRENT_RUNNING
    global PENDING_SCRAPERS
    global STOPPED_SCRAPERS

    if request.method == "POST":

        data = json.loads(request.body)

        shipping_lines = data.get(
            "shipping_lines",
            []
        )

        stopped = []

        for line in shipping_lines:

            # MARK STOPPED
            STOPPED_SCRAPERS.add(line)

            # REMOVE FROM PENDING
            if line in PENDING_SCRAPERS:
                PENDING_SCRAPERS.remove(line)

            process = RUNNING_SCRAPERS.get(line)

            if process:

                try:

                    process.kill()

                    stopped.append(line)

                    del RUNNING_SCRAPERS[line]

                    if line == CURRENT_RUNNING:
                        CURRENT_RUNNING = None

                except Exception as e:

                    print(e)

        return JsonResponse({
            "status": "success",
            "msg": f"Stopped: {', '.join(shipping_lines)}"
        })


# ================= STOP ALL =================

@csrf_exempt
def stop_all_scrapers(request):

    global RUNNING_SCRAPERS
    global CURRENT_RUNNING
    global PENDING_SCRAPERS
    global STOPPED_SCRAPERS

    STOPPED_SCRAPERS.update(PENDING_SCRAPERS)

    for line in list(RUNNING_SCRAPERS.keys()):

        STOPPED_SCRAPERS.add(line)

    stopped = []

    for line, process in list(RUNNING_SCRAPERS.items()):

        try:

            process.kill()

            stopped.append(line)

        except Exception as e:

            print(e)

    RUNNING_SCRAPERS.clear()

    PENDING_SCRAPERS.clear()

    CURRENT_RUNNING = None

    return JsonResponse({
        "status": "success",
        "msg": "All Scrapers Stopped"
    })





#****************************************************************

TRACK_DIR = os.path.join(BASE_DIR, "trackers")

TRACKERS = {
    "HAPAG": os.path.join(TRACK_DIR, "hapag_tracker.py"),
    "MAERSK": os.path.join(TRACK_DIR, "maersk_tracker.py"),
    "CMACGM": os.path.join(TRACK_DIR, "cmacgm_tracker.py")
}

# 🌐 TRACKER PAGE
def tracker_page(request):
    return render(request, "tracker.html")


# 📡 GET SHIPPING LINES (FOR DROPDOWN)
def get_shipping_lines(request):
    return JsonResponse({
        "shipping_lines": list(TRACKERS.keys())
    })


# ▶ RUN TRACKER
def run_tracker(request):
    if request.method == "POST":

        shipping_lines = request.POST.getlist("shipping_line[]")

        results = []

        for line in shipping_lines:

            if line in TRACKERS:
                script = TRACKERS[line]

                try:
                    result = subprocess.run(
                        ["python", script],
                        capture_output=True,
                        text=True
                    )

                    # 🔥 TERMINAL OUTPUT
                    print(f"\n===== {line} OUTPUT =====")
                    print(result.stdout)

                    if result.stderr:
                        print(f"===== {line} ERROR =====")
                        print(result.stderr)

                    # 📦 STORE FOR UI
                    results.append({
                        "line": line,
                        "output": result.stdout,
                        "error": result.stderr
                    })

                except Exception as e:
                    print(f"ERROR in {line}: {e}")

                    results.append({
                        "line": line,
                        "output": "",
                        "error": str(e)
                    })

        # ✅ SHOW RESULT PAGE
        return render(request, "tracker.html", {
            "results": results
        })

    return render(request, "tracker.html")