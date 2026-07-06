"""
Ablation Study
==============
این فایل رو در همون پوشه‌ای بذار که:
  - Main_Thesis__Final_version_.py
  - Data11.xlsx
هستن.

سه configuration مقایسه می‌شه:
  Config 1: Construction only  (Stages 1-3، بدون VNS)
  Config 2: Full algorithm     (Stages 1-3 + VNS) ← کد اصلی تو
  Config 3: Random init + VNS  (random grouping + VNS)

خروجی:
  - چاپ در terminal
  - ablation_results.xlsx
"""

from docplex.mp.model import Model
import sys
sys.path = [r'C:\Program Files\IBM\ILOG\CPLEX_Studio221\cplex\python\3.7\x64_win64'] + sys.path
import random
import numpy as np
import copy
import time
import openpyxl

# ── Seed ثابت ──
SEED = 42
random.seed(SEED)
np.random.seed(SEED)

# ══════════════════════════════════════════════════════
# خواندن داده از Excel
# ══════════════════════════════════════════════════════
workbook = openpyxl.load_workbook('Data11.xlsx')
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
                    row_data = []
                    for c in range(sc, sc + cols):
                        row_data.append(sheet.cell(row=r, column=c).value)
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

print(f"Instance: {orders1} orders | {num_picking_stations} stations | "
      f"group_cap={group_capacity} | max_iter={max_iterations}")
print("=" * 65)

# ══════════════════════════════════════════════════════
# توابع
# ══════════════════════════════════════════════════════

def create_order_groups(Sum_demand, orders, group_capacity):
    """Stage 1: Modified FFD"""
    order_groups = []
    remaining_orders = list(Sum_demand.copy())
    remaining_numbers = list(orders.copy())
    group_count = 0
    order_groups.append([])

    while len(remaining_orders) > 0:
        sorted_pairs = sorted(zip(remaining_orders, remaining_numbers), reverse=True)
        remaining_orders = [d for d, n in sorted_pairs]
        remaining_numbers = [n for d, n in sorted_pairs]
        order_groups[group_count].append(remaining_numbers[0])
        remaining_orders.pop(0)
        remaining_numbers.pop(0)

        if len(remaining_orders) > 0:
            if len(order_groups[group_count]) < group_capacity:
                sorted_pairs = sorted(zip(remaining_orders, remaining_numbers))
                remaining_orders = [d for d, n in sorted_pairs]
                remaining_numbers = [n for d, n in sorted_pairs]
                order_groups[group_count].append(remaining_numbers[0])
                remaining_orders.pop(0)
                remaining_numbers.pop(0)
            else:
                group_count += 1
                order_groups.append([])
                sorted_pairs = sorted(zip(remaining_orders, remaining_numbers))
                remaining_orders = [d for d, n in sorted_pairs]
                remaining_numbers = [n for d, n in sorted_pairs]
                order_groups[group_count].append(remaining_numbers[0])
                remaining_orders.pop(0)
                remaining_numbers.pop(0)

        if len(order_groups[group_count]) >= group_capacity and len(remaining_orders) > 0:
            group_count += 1
            order_groups.append([])

    return order_groups


def create_random_order_groups(orders, group_capacity):
    """Stage 1 جایگزین: Random grouping برای Config 3"""
    shuffled = list(orders)
    random.shuffle(shuffled)
    groups = []
    for i in range(0, len(shuffled), group_capacity):
        chunk = shuffled[i:i + group_capacity]
        if chunk:
            groups.append(chunk)
    return groups


def assigning_order_groups(order_groups, num_picking_stations, Demand):
    """Stage 2: Station assignment"""
    yy = np.zeros((len(order_groups), num_picking_stations, len(Demand)))
    active = np.zeros((len(order_groups), num_picking_stations, len(Demand)))
    ss, jj = 0, 0
    for g in range(len(order_groups)):
        yy[g, ss, jj] = 1
        active[g, ss, jj] = 1
        ss += 1
        if ss >= num_picking_stations:
            jj += 1
            ss = 0
    return yy, active, jj


def find_ss_jj(y, group, sequences, num_picking_stations):
    for ss in range(num_picking_stations):
        for jj in range(sequences):
            if y[group][ss][jj] == 1:
                return ss, jj
    return 0, 0


def swap_neighborhood(yy, sequences, num_picking_stations):
    new_y = yy.copy()
    group1, group2 = random.sample(range(len(new_y)), 2)
    ss1, jj1 = find_ss_jj(new_y, group1, sequences, num_picking_stations)
    ss2, jj2 = find_ss_jj(new_y, group2, sequences, num_picking_stations)
    new_y[group1][ss1][jj1] = 0
    new_y[group2][ss2][jj2] = 0
    new_y[group2][ss1][jj1] = 1
    new_y[group1][ss2][jj2] = 1
    return new_y


def perform_neighborhood_operation_1(order_groups):
    if len(order_groups) < 2:
        return order_groups
    g1, g2 = random.sample(range(len(order_groups)), 2)
    if order_groups[g1] and order_groups[g2]:
        i1 = random.randint(0, len(order_groups[g1]) - 1)
        i2 = random.randint(0, len(order_groups[g2]) - 1)
        order_groups[g1][i1], order_groups[g2][i2] = \
            order_groups[g2][i2], order_groups[g1][i1]
    return order_groups


def perform_neighborhood_operation_3(order_groups):
    if len(order_groups) < 2:
        return order_groups
    g1, g2 = random.sample(range(len(order_groups)), 2)
    order_groups[g1], order_groups[g2] = order_groups[g2], order_groups[g1]
    return order_groups


# evaluate_order_groups از کد اصلی import می‌شه
from Main_Thesis__Final_version_ import evaluate_order_groups


def run_vns(order_groups, yy, max_iterations, sequences, num_picking_stations):
    """اجرای VNS روی solution اولیه"""
    current_groups = copy.deepcopy(order_groups)
    current_cost = evaluate_order_groups(current_groups, yy)
    best_groups = copy.deepcopy(current_groups)
    best_cost = current_cost

    for _ in range(max_iterations):
        perform_neighborhood_operation_3(current_groups)
        perform_neighborhood_operation_1(current_groups)
        swap_neighborhood(yy, sequences, num_picking_stations)

        new_cost = evaluate_order_groups(current_groups, yy)
        if new_cost < current_cost:
            current_cost = new_cost
            best_groups = copy.deepcopy(current_groups)
            best_cost = current_cost

    return best_groups, best_cost


# ══════════════════════════════════════════════════════
# سه Configuration
# ══════════════════════════════════════════════════════

def config1_construction_only():
    t0 = time.time()
    groups = create_order_groups(Sum_demand, orders, group_capacity)
    yy, _, _ = assigning_order_groups(groups, num_picking_stations, Demand)
    cost = evaluate_order_groups(groups, yy)
    return cost, round(time.time() - t0, 2)


def config2_full_algorithm():
    t0 = time.time()
    groups = create_order_groups(Sum_demand, orders, group_capacity)
    yy, _, _ = assigning_order_groups(groups, num_picking_stations, Demand)
    best_groups, best_cost = run_vns(
        groups, yy, max_iterations, sequences, num_picking_stations)
    return best_cost, round(time.time() - t0, 2)


def config3_random_vns():
    t0 = time.time()
    groups = create_random_order_groups(orders, group_capacity)
    yy, _, _ = assigning_order_groups(groups, num_picking_stations, Demand)
    best_groups, best_cost = run_vns(
        groups, yy, max_iterations, sequences, num_picking_stations)
    return best_cost, round(time.time() - t0, 2)


# ══════════════════════════════════════════════════════
# اجرای Ablation Study
# ══════════════════════════════════════════════════════

def run_ablation(n_replications=5):
    configs = [
        ('Construction only (Stages 1-3)', config1_construction_only),
        ('Full algorithm (Stages 1-3 + VNS)', config2_full_algorithm),
        ('Random init + VNS', config3_random_vns),
    ]
    results = {name: {'costs': [], 'times': []} for name, _ in configs}

    for rep in range(n_replications):
        seed = SEED + rep
        random.seed(seed)
        np.random.seed(seed)
        print(f"\nReplication {rep+1}/{n_replications}  (seed={seed})")
        print("-" * 65)

        for name, func in configs:
            print(f"  {name}...", end=' ', flush=True)
            cost, t = func()
            results[name]['costs'].append(cost)
            results[name]['times'].append(t)
            print(f"cost={cost:.4f}  time={t}s")

    # ── نتایج ──
    print(f"\n\n{'='*75}")
    print("ABLATION STUDY RESULTS")
    print(f"{'='*75}")
    print(f"{'Configuration':<38} {'Mean':>8} {'Std':>8} {'Time':>8} {'Improv%':>9}")
    print("-" * 75)

    baseline = np.mean(results['Construction only (Stages 1-3)']['costs'])

    for name, _ in configs:
        data = results[name]
        mean_c = np.mean(data['costs'])
        std_c  = np.std(data['costs'])
        mean_t = np.mean(data['times'])
        improv = (baseline - mean_c) / baseline * 100 if baseline > 0 else 0
        print(f"{name:<38} {mean_c:>8.4f} {std_c:>8.4f} {mean_t:>8.2f} {improv:>8.1f}%")

    print("=" * 75)
    print(f"Note: Improv% = improvement over Construction-only baseline")
    print(f"Seeds: {SEED} to {SEED + n_replications - 1}")

    # ── ذخیره Excel ──
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Ablation Results"
    ws.append(["Configuration", "Replication", "Seed", "Cost", "Time (s)"])
    for name, _ in configs:
        for rep, (c, t) in enumerate(zip(results[name]['costs'], results[name]['times'])):
            ws.append([name, rep + 1, SEED + rep, round(c, 4), round(t, 2)])

    ws.append([])
    ws.append(["Summary"])
    ws.append(["Configuration", "Mean Cost", "Std Cost", "Mean Time (s)", "Improvement (%)"])
    for name, _ in configs:
        data = results[name]
        mean_c = np.mean(data['costs'])
        std_c  = np.std(data['costs'])
        mean_t = np.mean(data['times'])
        improv = (baseline - mean_c) / baseline * 100 if baseline > 0 else 0
        ws.append([name, round(mean_c,4), round(std_c,4), round(mean_t,2), round(improv,2)])

    wb.save("ablation_results.xlsx")
    print("\n✅ Results saved to ablation_results.xlsx")
    return results


if __name__ == "__main__":
    # اگه وقت کم داری n_replications=3 بذار
    run_ablation(n_replications=5)
