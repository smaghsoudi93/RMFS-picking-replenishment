# Per-neighborhood ablation (Reviewer comment R1-23).
"""
Table: Per-Neighborhood Contribution (Reviewer R1-23)
========================================================
این فایل رو در همون پوشه‌ای بذار که RMFS_main.py و instance موردنظر هستن.

⚠️ نسخه‌ی اصلاح‌شده: نسخه‌ی قبلی این فایل واقعاً فقط ۲ neighborhood
مجزا رو تست می‌کرد، نه ۳ تا -- "Neighborhood 1" (station/sequence
exchange از طریق order_groups) و چیزی که در RMFS_main.py با
swap_neighborhood پیاده شده بود (station/sequence exchange از طریق yy)
از نظر ریاضی دقیقاً یک نتیجه تولید می‌کردن (دو مکانیزم برای یک move).
نتیجه سوم (item بین wave، طبق متن Algorithm 4 مقاله) اصلاً در کد وجود
نداشت. الان به RMFS_main.py اضافه شده (تابع
perform_neighborhood_operation_wave_item + پارامتر wave_item_swaps در
evaluate_order_groups، به عنوان k=4 در variable_neighborhood_search -
کاملاً opt-in، پیش‌فرض active_neighborhoods=(1,2,3) دست‌نخورده مونده،
پس Table 3/6/7/Figure4 که قبلاً اجرا کردید همچنان معتبرن و نیازی به
اجرای مجدد ندارن).

کاری که این نسخه می‌کنه: روی همون instance، ۵ کانفیگ رو مقایسه می‌کنه
با ۳ neighborhood واقعاً مجزا:
  1. Construction only (بدون VNS، baseline)
  2. VNS با فقط Neighborhood "Order exchange" (k=2 → یک order بین دو گروه)
  3. VNS با فقط Neighborhood "Station/sequence exchange" (k=3 → swap_neighborhood)
  4. VNS با فقط Neighborhood "Item/wave exchange" (k=4 → تابع تازه اضافه‌شده)
  5. Full VNS (هر سه‌تا با هم، k=2,3,4 -- توجه: عمداً k=1 قدیمی رو کنار
     گذاشتیم چون همون station/sequence exchange رو (k=3) دوباره انجام
     می‌داد، صرفاً با مکانیزم دیگه -- نتیجه‌ی محاسباتی یکسان بود)

این دقیقاً جواب Comment R1-23 رو می‌ده: «سهم تک‌تک neighborhoodها» و
«توجیه ترتیب اجراشون» (هرکدوم بیشترین بهبود رو تنها داد، منطقی‌تره
که زودتر اجرا بشه).

استفاده:
  set RMFS_DATA_FILE=Data_XX.xlsx
  set NBHD_N_REPLICATIONS=3
  set NBHD_OUTPUT_FILE=neighborhood_ablation_XX.xlsx
  python run_neighborhood_ablation.py
"""

import sys
sys.path = [r'C:\Program Files\IBM\ILOG\CPLEX_Studio221\cplex\python\3.7\x64_win64'] + sys.path
import random
import numpy as np
import time
import os
import openpyxl

SEED = 42
random.seed(SEED)
np.random.seed(SEED)

N_REPLICATIONS = int(os.environ.get('NBHD_N_REPLICATIONS', '3'))
OUTPUT_FILE = os.environ.get('NBHD_OUTPUT_FILE', 'neighborhood_ablation_results.xlsx')
INSTANCE_LABEL = os.environ.get('NBHD_INSTANCE_LABEL', os.environ.get('RMFS_DATA_FILE', 'Data11.xlsx'))
FAIL_SENTINEL = 10000000

DATA_FILE = os.environ.get('RMFS_DATA_FILE', 'Data11.xlsx')
workbook = openpyxl.load_workbook(DATA_FILE)
sheet = workbook.active

def read_matrix(sheet, name):
    for row in range(1, sheet.max_row + 1):
        for col in range(1, sheet.max_column + 1):
            if sheet.cell(row=row, column=col).value == name:
                sr, sc = row + 1, col
                rows, cols = 0, 0
                while sheet.cell(row=sr + rows, column=sc).value is not None:
                    rows += 1
                while sheet.cell(row=sr, column=sc + cols).value is not None:
                    cols += 1
                data = []
                for r in range(sr, sr + rows):
                    row_data = [sheet.cell(row=r, column=c).value for c in range(sc, sc + cols)]
                    data.append(row_data)
                return data
    return None

def read_scalar(sheet, name):
    for row in range(1, sheet.max_row + 1):
        for col in range(1, sheet.max_column + 1):
            if sheet.cell(row=row, column=col).value == name:
                return sheet.cell(row=row + 1, column=col).value
    return None

Demand               = read_matrix(sheet, 'Demand')
Sum_demand           = np.sum(Demand, axis=1)
orders1              = read_scalar(sheet, 'orders1')
orders               = np.arange(orders1)
num_picking_stations = read_scalar(sheet, 'num_picking_stations')
group_capacity       = read_scalar(sheet, 'group_capacity')
max_iterations       = read_scalar(sheet, 'max_iterations')
sequences            = read_scalar(sheet, 'sequences')
items                = read_scalar(sheet, 'items')
waves                = read_scalar(sheet, 'waves')

from RMFS_main import (
    create_order_groups, assigning_order_groups,
    evaluate_order_groups, variable_neighborhood_search,
)

print(f"Instance label: {INSTANCE_LABEL}")
print(f"Instance: {orders1} orders | max_iterations={max_iterations}")
print("=" * 65)

CONFIGS = [
    ("Construction only", None),
    ("VNS: Order exchange only", (2,)),
    ("VNS: Station/sequence exchange only", (3,)),
    ("VNS: Item/wave exchange only", (4,)),
    ("Full VNS (all 3 distinct neighborhoods)", (2, 3, 4)),
]


def run_config(active_nbhd):
    t0 = time.time()
    if active_nbhd is None:
        groups = create_order_groups(Sum_demand, orders, group_capacity)
        yy, _, _ = assigning_order_groups(groups, num_picking_stations, Demand)
        cost = evaluate_order_groups(groups, yy)
    else:
        _, cost, _ = variable_neighborhood_search(
            Sum_demand, orders, group_capacity, max_iterations,
            num_picking_stations, Demand, sequences,
            active_neighborhoods=active_nbhd,
            items=items, waves=waves)
    return cost, time.time() - t0


results = {name: {'costs': [], 'times': []} for name, _ in CONFIGS}

for rep in range(N_REPLICATIONS):
    seed = SEED + rep
    print(f"\nReplication {rep+1}/{N_REPLICATIONS}  (seed={seed})")
    print("-" * 65)
    for name, active_nbhd in CONFIGS:
        random.seed(seed)
        np.random.seed(seed)
        print(f"  {name}...", end=' ', flush=True)
        cost, t = run_config(active_nbhd)
        results[name]['costs'].append(cost)
        results[name]['times'].append(t)
        status = "" if cost < FAIL_SENTINEL else "  [FAILED]"
        print(f"cost={cost:.4f}  time={t:.1f}s{status}")

print(f"\n\n{'='*80}")
print(f"PER-NEIGHBORHOOD ABLATION -- {INSTANCE_LABEL}")
print(f"{'='*80}")
print(f"{'Configuration':<55} {'Mean(ok)':>9} {'Std(ok)':>9} {'Success':>8}")
print("-" * 80)

stats = {}
baseline = None
for name, _ in CONFIGS:
    data = results[name]
    ok = [c for c in data['costs'] if c < FAIL_SENTINEL]
    mean_c = np.mean(ok) if ok else None
    std_c = np.std(ok) if ok else None
    stats[name] = {'mean_c': mean_c, 'std_c': std_c, 'n_ok': len(ok)}
    if name == "Construction only":
        baseline = mean_c

for name, _ in CONFIGS:
    s = stats[name]
    mean_str = f"{s['mean_c']:.4f}" if s['mean_c'] is not None else "N/A"
    std_str = f"{s['std_c']:.4f}" if s['std_c'] is not None else "N/A"
    succ_str = f"{s['n_ok']}/{N_REPLICATIONS}"
    print(f"{name:<55} {mean_str:>9} {std_str:>9} {succ_str:>8}")

print("-" * 80)
print("Improvement over Construction-only baseline:")
for name, _ in CONFIGS:
    if name == "Construction only":
        continue
    s = stats[name]
    if baseline and s['mean_c'] is not None:
        improv = (baseline - s['mean_c']) / baseline * 100
        print(f"  {name}: {improv:.2f}%")
print("=" * 80)

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Neighborhood Ablation"
ws.append(["Instance", INSTANCE_LABEL])
ws.append(["Configuration", "Replication", "Seed", "Cost", "Time (s)", "Failed?"])
for name, _ in CONFIGS:
    for rep, (c, t) in enumerate(zip(results[name]['costs'], results[name]['times'])):
        ws.append([name, rep + 1, SEED + rep,
                   round(c, 4) if c < FAIL_SENTINEL else "FAILED",
                   round(t, 2), c >= FAIL_SENTINEL])

ws.append([])
ws.append(["Summary (means over successful replications only)"])
ws.append(["Configuration", "Mean Cost (ok)", "Std Cost (ok)", "Successes", "Total Reps",
           "Improvement vs Construction-only (%)"])
for name, _ in CONFIGS:
    s = stats[name]
    improv = None
    if name != "Construction only" and baseline and s['mean_c'] is not None:
        improv = (baseline - s['mean_c']) / baseline * 100
    ws.append([name,
               round(s['mean_c'], 4) if s['mean_c'] is not None else "N/A",
               round(s['std_c'], 4) if s['std_c'] is not None else "N/A",
               s['n_ok'], N_REPLICATIONS,
               round(improv, 2) if improv is not None else ("—" if name == "Construction only" else "N/A")])

wb.save(OUTPUT_FILE)
print(f"\n✅ Results saved to {OUTPUT_FILE}")
