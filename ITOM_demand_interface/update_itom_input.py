# Writes the new Demand and Scrap data (Demand.csv, Scrap.csv from demand_and_scrap_availability.py)
# into the ITOM model input file and saves the result as a new file. Everything else is copied unchanged.

from copy import copy
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

# Settings ######

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "_model_data"

INPUT_XLSX = BASE_DIR / "ITOM_Steel_TRANSIENCE_baseline0_input.xlsx"
OUTPUT_XLSX = BASE_DIR / "ITOM_Steel_TRANSIENCE_baseline0_input_updated.xlsx"

DEMAND_SHEET = "Demand"
SCRAP_SHEET = "Total...AnnualActivityUpLimit"

# Region names used in the ITOM input file
DEMAND_REGION = "Europe_Demand"
SCRAP_REGION = "energy_center"

SCRAP_PREFIX = "import_scrap"


def read_csv2(path):
    return pd.read_csv(path, sep=";", decimal=",")


def replace_rows(ws, rows):
    """Replaces all data rows (below the header) of a sheet, keeping the cell style of the first data row."""
    styles = [copy(ws.cell(row=2, column=c)._style) for c in range(1, len(rows[0]) + 1)]
    ws.delete_rows(2, max(ws.max_row - 1, 1))
    for r, row in enumerate(rows, start=2):
        for c, value in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=value)
            cell._style = copy(styles[c - 1])


# Load ######

print("Loading", INPUT_XLSX.name)
wb = load_workbook(INPUT_XLSX)

# 1. Demand: existing data is overwritten ######

demand = read_csv2(DATA_DIR / "Demand.csv")
demand["Region"] = DEMAND_REGION
demand = demand.sort_values(["Product", "Year"], kind="stable")

ws = wb[DEMAND_SHEET]
header = [c.value for c in ws[1]][:4]
assert header == ["REGION", "PRODUCT", "YEAR", "SpecifiedAnnualDemand"], header

replace_rows(ws, demand[["Region", "Product", "Year", "SpecifiedAnnualDemand"]].values.tolist())
print(f"Demand: {len(demand)} rows written")

# 2. Scrap: existing scrap availabilities are removed, all other rows are kept ######

scrap = read_csv2(DATA_DIR / "Scrap.csv")
scrap["Region"] = SCRAP_REGION
new_rows = scrap[["Region", "Technology", "Year", "TotalTechnologyAnnualActivityUpperLimit"]].values.tolist()

ws = wb[SCRAP_SHEET]
header = [c.value for c in ws[1]][:4]
assert header == ["REGION", "TECHNOLOGY", "YEAR", "TotalTechnologyAnnualActivityUpperLimit"], header

old_rows = [[c.value for c in row[:4]] for row in ws.iter_rows(min_row=2) if row[0].value is not None]

# The new scrap rows take the place of the first removed scrap row
first_scrap = next((i for i, r in enumerate(old_rows) if str(r[1]).startswith(SCRAP_PREFIX)), len(old_rows))
kept_before = [r for r in old_rows[:first_scrap]]
kept_after = [r for r in old_rows[first_scrap:] if not str(r[1]).startswith(SCRAP_PREFIX)]
n_removed = len(old_rows) - len(kept_before) - len(kept_after)

replace_rows(ws, kept_before + new_rows + kept_after)
print(f"Scrap: {n_removed} old rows removed, {len(new_rows)} rows written, {len(kept_before) + len(kept_after)} other rows kept")

# Save ######

wb.save(OUTPUT_XLSX)
print("Saved", OUTPUT_XLSX.name)
