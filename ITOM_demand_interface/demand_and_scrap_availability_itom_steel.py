# ITOM steel: Python version of "4_Demand and Scrap availability_october26-update.R"
# Reads the MFA output files from output_to_itom and writes itom_steel_Demand.csv and itom_steel_Scrap.csv for the ITOM steel model.

from pathlib import Path

import numpy as np
import pandas as pd

# Settings ######

BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "output_to_itom"
OUTPUT_DIR = BASE_DIR / "_model_data"

DEMAND_FILE = INPUT_DIR / "FinalDemand.csv"
SCRAP_FILE = INPUT_DIR / "waste_management__available_scrap_sysenv.csv"
PRODUCTION_SCRAP_FILE = INPUT_DIR / "steel_goods_manufacturing__sysenv.csv"
MAPPING_FILE = BASE_DIR / "mfa_itom_steel_product_mapping.xlsx"
MAPPING_SHEET = "products_mfa_itom"

FIRST_YEAR = 2018  # earlier years are removed from both input files
LAST_YEAR = 2050
INPUT_YEARS = list(range(FIRST_YEAR, LAST_YEAR + 1))
# Model years: each one gets the mean of the 5-year window centred on it (year-2 ... year+2)
YEAR = list(range(2020, LAST_YEAR + 1, 5))
WINDOW_HALF_WIDTH = 2
# The model input holds the total over the 5-year period, not the annual mean
YEARS_PER_PERIOD = 5

# Region name used in the model for "EU27+1" (energy_center_geography$Region[1] in the R version)
MODEL_REGION = "Europe"
SOURCE_REGION = "EU27+1"

# The R version divided the MFA output by 1000 (kt -> Mt)
UNIT_DIVISOR = 1000

# Decimal places of the values in the output files
DECIMALS = 3

TECH_NAMES = ["import_scrap_0", "import_scrap_01", "import_scrap_02",
              "import_scrap_03", "import_scrap_04", "import_scrap_05"]


def read_input(path):
    """Reads an input file and removes the years before FIRST_YEAR."""
    df = pd.read_csv(path)
    return df[df["time"] >= FIRST_YEAR].copy()


def five_year_mean(df, keys, value):
    """Mean of the annual values in the 5-year window centred on each model year.

    The window is cut off at the end of the data (2050 uses 2048-2050).
    """
    windows = []
    for year in YEAR:
        in_window = df[(df["Year"] >= year - WINDOW_HALF_WIDTH) & (df["Year"] <= year + WINDOW_HALF_WIDTH)]
        windows.append(in_window.groupby(keys, as_index=False)[value].mean().assign(Year=year))
    return pd.concat(windows, ignore_index=True)


def write_csv2(df, path):
    """Equivalent of R's write.csv2 (sep ';', decimal ',')."""
    path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(path, sep=";", decimal=",", index=False)


# 1. Creating the demand data ######

print("Processing demand data")

demand_data = read_input(DEMAND_FILE)
demand_data = demand_data[demand_data["element"] == "All"]

# Map MFA sector/product to the ITOM demand product (estimated copper tolerance only)
keys_demand_products = pd.read_excel(MAPPING_FILE, sheet_name=MAPPING_SHEET)[
    ["sectors", "products", "Product_EST"]]
demand_data = demand_data.merge(
    keys_demand_products, left_on=["sector", "product"], right_on=["sectors", "products"], how="left")

# Sector/product combinations missing in the mapping sheet are only accepted if they carry no demand
unmapped = demand_data[demand_data["Product_EST"].isna()]
if unmapped["value"].abs().sum() > 0:
    missing = unmapped[unmapped["value"] != 0][["sector", "product"]].drop_duplicates()
    raise ValueError(f"No ITOM product in the mapping sheet for:\n{missing.to_string(index=False)}")
demand_data = demand_data.dropna(subset=["Product_EST"])

demand_data = demand_data.assign(
    Region=demand_data["region"].replace({SOURCE_REGION: MODEL_REGION}),
    Product=demand_data["Product_EST"],
    Year=demand_data["time"],
    SpecifiedAnnualDemand=demand_data["value"] / UNIT_DIVISOR,
)

Demand = (demand_data
          .groupby(["Region", "Product", "Year"], as_index=False)["SpecifiedAnnualDemand"].sum())
Demand["Product"] = "Final_" + Demand["Product"]
Demand = five_year_mean(Demand, ["Region", "Product"], "SpecifiedAnnualDemand")
Demand = Demand[["Region", "Product", "Year", "SpecifiedAnnualDemand"]].dropna()
Demand["SpecifiedAnnualDemand"] = (Demand["SpecifiedAnnualDemand"] * YEARS_PER_PERIOD).round(DECIMALS)

write_csv2(Demand, OUTPUT_DIR / "itom_steel_Demand.csv")
print("Demand data has been processed and saved")

# 2 Scrap Data ######
# 2.1 Proportional allocation of scrap to copper categories via Cu concentration weights ######

print("Processing scrap availability")

raw_scrap = read_input(SCRAP_FILE).rename(columns={
    "region": "regions", "time": "years", "waste_category": "wastes"})
keys = ["regions", "years", "wastes"]

scrap_all = raw_scrap[raw_scrap["element"] == "All"].rename(columns={"value": "All"})[keys + ["All"]]
scrap_cu = raw_scrap[raw_scrap["element"] == "Cu"].rename(columns={"value": "Cu"})[keys + ["Cu"]]

# Linear interpolation weights from the Cu concentration (per mille) between the categories
weights = scrap_all.merge(scrap_cu, on=keys, how="left")
weights["conc"] = np.where(weights["All"] > 0, weights["Cu"] / weights["All"] * 1000, 0)
weights["conc_clamped"] = weights["conc"].clip(0, 5)
weights["k"] = np.minimum(np.floor(weights["conc_clamped"]), 4).astype(int)
weights["w_upper"] = weights["conc_clamped"] - weights["k"]
weights["w_lower"] = 1 - weights["w_upper"]

tech = np.array(TECH_NAMES)
lower = weights.assign(scrap_category=tech[weights["k"]], weight=weights["w_lower"])
upper = weights.assign(scrap_category=tech[weights["k"] + 1], weight=weights["w_upper"])
scrap_data = pd.concat([lower, upper])
scrap_data = scrap_data[scrap_data["weight"] > 0].copy()
scrap_data["All"] = scrap_data["All"] * scrap_data["weight"]

# 2.2 Production scrap (steel goods manufacturing) goes into import_scrap_0 ######

production_scrap = read_input(PRODUCTION_SCRAP_FILE)
production_scrap = (production_scrap[production_scrap["element"] == "All"]
                    .groupby(["region", "time"], as_index=False)["value"].sum()
                    .rename(columns={"region": "regions", "time": "years", "value": "All"})
                    .assign(scrap_category="import_scrap_0"))

scrap_data = pd.concat([scrap_data, production_scrap], ignore_index=True)

# 2.3 Format output table ######

scrap_data = (scrap_data
              .groupby(["regions", "years", "scrap_category"], as_index=False)["All"].sum()
              .rename(columns={"regions": "Region", "years": "Year", "scrap_category": "Technology",
                               "All": "TotalTechnologyAnnualActivityUpperLimit"}))
scrap_data["Region"] = scrap_data["Region"].replace({SOURCE_REGION: MODEL_REGION})
scrap_data["TotalTechnologyAnnualActivityUpperLimit"] /= UNIT_DIVISOR

# 2.4 Full grid completion across all input years and all scrap categories ######

grid = pd.MultiIndex.from_product(
    [sorted(scrap_data["Region"].unique()), INPUT_YEARS, TECH_NAMES],
    names=["Region", "Year", "Technology"]).to_frame(index=False)

Scrap = grid.merge(scrap_data, on=["Region", "Year", "Technology"], how="left")
Scrap["TotalTechnologyAnnualActivityUpperLimit"] = Scrap["TotalTechnologyAnnualActivityUpperLimit"].fillna(0)

# Missing years/categories count as 0 in the mean, so the average is taken on the completed grid
Scrap = five_year_mean(Scrap, ["Region", "Technology"], "TotalTechnologyAnnualActivityUpperLimit")
Scrap["TotalTechnologyAnnualActivityUpperLimit"] = (
    Scrap["TotalTechnologyAnnualActivityUpperLimit"] * YEARS_PER_PERIOD).round(DECIMALS)
Scrap = Scrap[["Region", "Year", "Technology", "TotalTechnologyAnnualActivityUpperLimit"]]

Scrap = Scrap[Scrap["Technology"] != "import_scrap_05"]

write_csv2(Scrap, OUTPUT_DIR / "itom_steel_Scrap.csv")
print("Scrap availability has been processed and saved")
