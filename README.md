# RMFS Picking-Replenishment Optimization

Code accompanying the paper *"[Paper title]"* (under Major Revision), which
studies joint optimization of order picking and replenishment in
Robotic Mobile Fulfillment Systems (RMFS). This repository contains an
exact MIP model (CPLEX/docplex) and a four-stage heuristic (grouping →
station/sequence assignment → wave planning → variable neighborhood
search) for the same problem.

## Requirements

- Python 3.7+
- IBM ILOG CPLEX Optimization Studio (with a valid license) and its
  `docplex` Python API
- `openpyxl`, `numpy`, `scipy`

The main script currently hard-codes the CPLEX Python API path for
Windows:
```python
sys.path = [r'C:\Program Files\IBM\ILOG\CPLEX_Studio221\cplex\python\3.7\x64_win64'] + sys.path
```
Update this path to match your local CPLEX installation.

## Repository structure

| File | Purpose |
|---|---|
| `RMFS_main.py` | Main file — MIP model + all core functions (`evaluate_order_groups`, `variable_neighborhood_search`, `final_polish`, `evaluate_order_groups_sequential`, `evaluate_order_groups_rulebased`, etc.). Every other script imports from this file. |
| `Standard_deviation.py` | Table 3 / Table 3b — main results on the paper's 12 core instances |
| `run_ffd_multi.py` | Table 4 — FFD grouping strategy comparison |
| `Stage3_variance.py` | Table 5 — Stage 3 (wave-assignment) variance across random draws |
| `run_ablation.py` / `run_ablation_multi.py` | Table 6 — construction vs. VNS ablation |
| `run_sa.py` | Table 7 — VNS vs. Simulated Annealing |
| `run_neighborhood_ablation.py` | Per-neighborhood contribution ablation (order exchange / station-sequence exchange / item-wave exchange) |
| `Data*.xlsx` | Problem instances (orders, demand, capacities, etc.) |

## Usage

Each `run_*.py` script reads its instance and configuration from
environment variables, then imports the required functions from
`RMFS_main.py`. Example (Windows Command Prompt/PowerShell):

```
set RMFS_DATA_FILE=Data4.xlsx
set ABLATION_OUTPUT_FILE=ablation_results_Data4.xlsx
python run_ablation.py
```

**Note:** run each instance individually rather than through the
`run_*_multi.py` batch wrappers where possible — the multi-instance
subprocess wrappers can hang or buffer unpredictably on Windows. Run
scripts from Command Prompt/PowerShell, not IDLE, for the same reason.

Each script writes its results to an `.xlsx` file (name configurable
via the corresponding `*_OUTPUT_FILE` environment variable).

## Reproducibility

All reported experiments use a fixed random seed (`seed = 42`, with
seeds 42–46 for the 5 Stage-3 replications). CPLEX is run with 4
threads, a 5% MIP gap, and an instance-size-scaled time limit (see
`SOLVE_TIME_LIMIT_SECONDS` / `SOLVE_MIPGAP` at the top of
`RMFS_main.py`).

## Citation

If you use this code, please cite:

```
[citation to be added upon publication]
```

## License

[Add license here, e.g., MIT]
