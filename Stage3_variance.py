# Table 5: Stage 3 replication variance.
"""
Table 5 -- Stage 3 Replication Variance
=========================================
این فایل رو در همون پوشه‌ای بذار که RMFS_main.py و instance موردنظر هستن.

کاری که می‌کنه:
  - Stage 1 (گروه‌بندی FFD) و Stage 2 (تخصیص ایستگاه) رو فقط یک‌بار
    اجرا می‌کنه و ثابت نگه می‌داره.
  - فقط Stage 3 (heuristic موج‌بندی تصادفی + حل MIP، یعنی خود
    evaluate_order_groups) رو با seedهای مختلف N بار تکرار می‌کنه.
  - این دقیقاً همون چیزیه که Table 5 در پاسخ به Reviewer (کامنت درباره‌ی
    randomization توجیه‌نشده‌ی Stage 3) نیاز داره: نشون می‌ده وقتی
    گروه‌بندی/ایستگاه‌ها ثابته، فقط تصادفی‌بودن موج‌بندی چقدر روی جواب
    نهایی اثر می‌ذاره.

استفاده:
  set RMFS_DATA_FILE=Data_XX.xlsx
  set TABLE5_N_REPLICATIONS=5
  set TABLE5_OUTPUT_FILE=table5_XX.xlsx
  python run_table5.py

خروجی: table5_<label>.xlsx (یا هر اسمی که با TABLE5_OUTPUT_FILE بدید)
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

N_REPLICATIONS = int(os.environ.get('TABLE5_N_REPLICATIONS', '5'))
OUTPUT_FILE = os.environ.get('TABLE5_OUTPUT_FILE', 'table5_results.xlsx')
INSTANCE_LABEL = os.environ.get('TABLE5_INSTANCE_LABEL', os.environ.get('RMFS_DATA_FILE', 'Data11.xlsx'))
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

from RMFS_main import create_order_groups, assigning_order_groups, evaluate_order_groups

print(f"Instance label: {INSTANCE_LABEL}")
print(f"Instance: {orders1} orders | {num_picking_stations} picking stations | "
      f"group_capacity={group_capacity}")
print("=" * 65)

# ── Stage 1 + Stage 2 اجرا و برای همیشه ثابت می‌شن ──
fixed_groups = create_order_groups(Sum_demand, orders, group_capacity)
fixed_yy, _, _ = assigning_order_groups(fixed_groups, num_picking_stations, Demand)
print(f"Fixed grouping (Stage 1+2, will NOT change across replications): {fixed_groups}")
print("=" * 65)

costs, times = [], []
for rep in range(N_REPLICATIONS):
    seed = SEED + rep
    random.seed(seed)
    np.random.seed(seed)
    print(f"\nReplication {rep+1}/{N_REPLICATIONS}  (seed={seed}) -- Stage 3 only")
    t0 = time.time()
    cost = evaluate_order_groups(fixed_groups, fixed_yy)
    dt = time.time() - t0
    costs.append(cost)
    times.append(dt)
    status = "" if cost < FAIL_SENTINEL else "  [FAILED]"
    print(f"  -> objective={cost:.4f}  time={dt:.1f}s{status}")

ok_costs = [c for c in costs if c < FAIL_SENTINEL]
n_ok = len(ok_costs)
mean_c = np.mean(ok_costs) if ok_costs else None
std_c = np.std(ok_costs) if ok_costs else None
cv = (std_c / mean_c * 100) if (mean_c and mean_c > 0) else None

print(f"\n\n{'='*65}")
print(f"TABLE 5 RESULT -- {INSTANCE_LABEL}")
print(f"{'='*65}")
print(f"Successful replications: {n_ok}/{N_REPLICATIONS}")
if mean_c is not None:
    print(f"Mean objective: {mean_c:.4f}  |  Std. Dev.: {std_c:.4f}"
          + (f"  |  Coefficient of variation: {cv:.2f}%" if cv is not None else ""))
else:
    print("⚠️  ALL replications failed.")
print("=" * 65)

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Table 5"
ws.append(["Instance", INSTANCE_LABEL])
ws.append(["Fixed grouping (Stage 1+2)", str(fixed_groups)])
ws.append([])
header = ["Instance"] + [f"Rep {i+1}" for i in range(N_REPLICATIONS)] + ["Mean", "Std. Dev."]
ws.append(header)
row = [INSTANCE_LABEL] + [round(c, 4) if c < FAIL_SENTINEL else "FAILED" for c in costs]
row += [round(mean_c, 4) if mean_c is not None else "N/A",
        round(std_c, 4) if std_c is not None else "N/A"]
ws.append(row)

ws.append([])
ws.append(["Replication", "Seed", "Objective", "Time (s)", "Failed?"])
for rep, (c, t) in enumerate(zip(costs, times)):
    ws.append([rep + 1, SEED + rep, round(c, 4) if c < FAIL_SENTINEL else "FAILED",
               round(t, 2), c >= FAIL_SENTINEL])

wb.save(OUTPUT_FILE)
print(f"\n✅ Results saved to {OUTPUT_FILE}")
