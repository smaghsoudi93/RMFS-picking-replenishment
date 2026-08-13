# RMFS Picking-Replenishment Optimization

Code accompanying the paper *"A comprehensive approach to joint picking and
replenishment optimization in robotic mobile fulfillment systems"*, which
studies joint optimization of order picking and replenishment in Robotic
Mobile Fulfillment Systems (RMFS).

The repository contains two distinct implementations:

- **`exact_milp.py`** — the standalone exact MILP of Section 3, Constraints
  (3.1)–(3.26), with every decision left free for CPLEX. This is the model
  reported in the CPLEX columns of Table 3.
- **`RMFS_main.py`** — the four-stage heuristic of Section 4 (order grouping →
  station/sequence assignment → wave planning → variable neighborhood search).
  The MIP built inside this file is the *reduced* sub-problem of Section 4.3,
  solved with grouping and station assignment already fixed by Stages 1–2. It
  is not the exact model.

## Requirements

- Python 3.7+
- IBM ILOG CPLEX Optimization Studio 22.1 with a valid license, plus its
  `docplex` Python API
- See `requirements.txt` for exact package versions:
  `python -m pip install -r requirements.txt`

If `import cplex` fails, point the `CPLEX_PYTHON_API` environment variable at
the folder holding the CPLEX Python API:

```
# Windows
set CPLEX_PYTHON_API=C:\Program Files\IBM\ILOG\CPLEX_Studio221\cplex\python\3.7\x64_win64
# Linux
export CPLEX_PYTHON_API=/opt/ibm/ILOG/CPLEX_Studio221/cplex/python/3.7/x86-64_linux
# macOS
export CPLEX_PYTHON_API=/Applications/CPLEX_Studio221/cplex/python/3.7/x86-64_osx
```

The CPLEX Community Edition bundled with `pip install cplex` is limited to
1000 variables and 1000 constraints, which covers only the smallest instances.
A full CPLEX installation (free for academics via the IBM Academic Initiative)
is required to reproduce the reported results.

## Repository structure

| File | Purpose |
|---|---|
| `exact_milp.py` | Standalone exact MILP — CPLEX columns of Table 3 |
| `RMFS_main.py` | Four-stage heuristic: reduced MIP, `evaluate_order_groups`, `variable_neighborhood_search`, `final_polish`. Every other script imports from this file. |
| `Standard_deviation.py` | Table 3 / Table 3b — heuristic results and standard deviations |
| `run_ffd_single.py` | Table 4 — grouping-strategy comparison on one instance |
| `run_ffd_multi.py` | Table 4 — batch wrapper over several instances |
| `Stage3_variance.py` | Table 5 — Stage 3 variance across random draws |
| `ablation_study.py` | Table 6 — construction vs. VNS ablation |
| `run_sa.py` | Table 7 — VNS vs. Simulated Annealing |
| `run_neighborhood_ablation.py` | Table 8 — per-neighborhood contribution |
| `Data*.xlsx` | Problem instances (orders, demand, capacities, wave arrivals) |

`Data1.xlsx` … `Data12.xlsx` correspond to the 12 rows of Table 3 in order
(6, 8, 15, 20, 30, 40, 50, 60, 70, 80, 90, 100 orders).

## Usage

Each script reads its instance and configuration from environment variables.
Example (Windows Command Prompt/PowerShell):

```
set RMFS_DATA_FILE=Data4.xlsx
set ABLATION_OUTPUT_FILE=ablation_results_Data4.xlsx
python ablation_study.py
```

Example (macOS/Linux):

```
RMFS_DATA_FILE=Data4.xlsx ABLATION_OUTPUT_FILE=ablation_Data4.xlsx python ablation_study.py
```

To reproduce the exact-MILP benchmark for one instance:

```
RMFS_DATA_FILE=Data4.xlsx EXACT_OUTPUT_FILE=exact_Data4.xlsx python exact_milp.py
```

Run each instance individually rather than through the `*_multi.py` batch
wrappers where possible — the subprocess wrappers can buffer unpredictably on
Windows. Run scripts from Command Prompt/PowerShell, not IDLE, for the same
reason.

Each script writes results to an `.xlsx` file, named via the corresponding
`*_OUTPUT_FILE` environment variable.

## The neighborhood structures

`variable_neighborhood_search` indexes its moves k = 1..4:

| k | Move |
|---|---|
| 1 | Exchange the contents of two groups, leaving the station assignment unchanged |
| 2 | Move one order between two groups (**N1** in the manuscript) |
| 3 | Exchange the station and sequence position of two groups (**N2**) |
| 4 | Exchange one item between two waves (**N3**) |

k=1 and k=3 are two mechanisms for the same move — both end in a different
assignment of groups to (station, sequence) pairs — so a configuration
containing both would count one neighborhood twice. The proposed VNS is
defined once, as `ACTIVE_NEIGHBORHOODS` in `RMFS_main.py`, and every script
reporting a headline result imports that constant.

## Sequence positions

The number of sequence positions modelled per picking station, |Sq|, is set by
`RMFS_SEQ_MODE`:

- `tight` (default) — |Sq| = number of order groups. This is the smallest
  valid bound, since at most every group could queue at one station, and it is
  the definition given in the notation table of Section 3.1.
- `legacy` — |Sq| = |O|, the looser bound used to produce the numbers in the
  published tables.

Both bounds are valid: surplus positions stay empty and Constraints
(3.23)–(3.24) force their start and completion times to zero, so the optimal
objective is unchanged. The tight bound simply builds a smaller model. Set
`RMFS_SEQ_MODE=legacy` to reproduce the published numbers exactly.

## Reproducibility

All reported experiments use a fixed random seed (`seed = 42`, with seeds
42–46 for the 5 Stage-3 replications). CPLEX is run with 4 threads and the
default branching strategy. `k_max` is 100 for all experiments, and Table 3
results are averaged over 10 independent runs.

## Citation

Archived release: https://doi.org/10.5281/zenodo.21225312
