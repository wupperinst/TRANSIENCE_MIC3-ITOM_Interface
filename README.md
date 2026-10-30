# TRANSIENCE ITOM–MIC3 Interface

Interface scripts that convert output data from different modules of the **MIC3 framework** into the input format of the **ITOM steel model** (`ITOM_Steel_TRANSIENCE_baseline0_input.xlsx`).

Each interface reads data from one MIC3 module and writes it into the relevant sheets of the ITOM input workbook. The scripts never modify the workbook they read; they always save a new file.

| Interface | MIC3 module | Data written to ITOM | Sector |
|---|---|---|---|
| [`ITOM_demand_interface`](ITOM_demand_interface) | MFA (material flow analysis) | Final steel demand and scrap availability | steel |
| [`ITOM_Energy_Interface`](ITOM_Energy_Interface) | OPEN-PROM (energy system model) | Energy carrier prices and biomass availability | steel |
| [`ITOM_ETS_Interface`](ITOM_ETS_Interface) | Carbon price projections (EU ETS) | CO₂ emissions penalty | all sectors |

### Sector naming convention

The demand and energy interfaces are specific to the steel sector. The sector is part of the file names as `itom_steel` (the ITOM model plus its sector): scripts are named like `update_itom_steel_input.py` or `demand_and_scrap_availability_itom_steel.py`, the config is `config_itom_steel.yaml`, and the intermediate files are `itom_steel_Demand.csv`, `itom_steel_Scrap.csv`, `itom_steel_open_prom_relative_changes.csv` and `mfa_itom_steel_product_mapping.xlsx`. Workflows for further industrial sectors will be added to the same folders under the same pattern (for example `update_itom_<sector>_input.py`, `config_itom_<sector>.yaml`, `itom_<sector>_Demand.csv`). The ETS interface works for all sectors and has no sector name.

## Workflow

The three interfaces are designed to be applied one after another. Each step takes the workbook produced by the previous one as its input:

```
ITOM_Steel_TRANSIENCE_baseline0_input.xlsx
        │  1. ITOM_demand_interface   (MFA output CSVs)
        ▼
ITOM_Steel_TRANSIENCE_baseline0_input_MFA-update.xlsx
        │  2. ITOM_Energy_Interface   (OPEN-PROM Excel export)
        ▼
ITOM_Steel_TRANSIENCE_baseline0_input_MFA-Energy-update.xlsx
        │  3. ITOM_ETS_Interface      (CarbonPrices.xlsx)
        ▼
ITOM_Steel_TRANSIENCE_baseline0_input_MFA-Energy-ETS_update.xlsx
```

The intermediate file names above are the ones configured in the scripts. The demand interface writes `..._input_updated.xlsx`; rename or copy it to `..._input_MFA-update.xlsx` (or adjust the paths in `config_itom_steel.yaml`) before running step 2.

The model runs on 5-year steps: **2020, 2025, …, 2050**.

## Setup

All three interfaces share one virtual environment in the repository root and one [`requirements.txt`](requirements.txt) (`pandas`, `numpy`, `openpyxl`, `pyyaml`, `pillow`). Python 3.13 is recommended. `pillow` makes `openpyxl` keep the images in the workbook's Info sheet.

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

With [uv](https://docs.astral.sh/uv/) the same setup is `uv venv && uv pip install -r requirements.txt`.

Activate the environment in the repository root, then run the scripts from inside their interface folder as shown below.

## 1. Demand and scrap availability (MFA)

[`ITOM_demand_interface/`](ITOM_demand_interface) is a Python port of the R script `4_Demand and Scrap availability_october26-update.R` (Süheyb Bilici & Alexander Kling). The R file is kept for reference only and is not executed.

It runs in two steps:

| Script | Role |
|---|---|
| `demand_and_scrap_availability_itom_steel.py` | Reads the MFA CSVs from `output_to_itom/` and writes `_model_data/itom_steel_Demand.csv` and `_model_data/itom_steel_Scrap.csv` |
| `update_itom_steel_input.py` | Writes those CSVs into the `Demand` and `Total...AnnualActivityUpLimit` sheets of the ITOM workbook and saves a new file |

Supporting files:

- `mfa_itom_steel_product_mapping.xlsx` maps MFA sector/product to the ITOM demand product (sheet `products_mfa_itom`).
- `output_to_itom/` holds the MFA output CSVs. `output_to_itom.zip` contains the same data.

What the processing does:

- Years before 2018 are dropped. Each model year is the mean of a 5-year window (year−2 to year+2), then multiplied by 5, because ITOM stores the total over the 5-year period.
- Scrap is split into `import_scrap_0` … `import_scrap_04` categories by copper concentration, using linear interpolation between neighbouring categories.
- Production scrap is added to `import_scrap_0`.

```bash
cd ITOM_demand_interface
python demand_and_scrap_availability_itom_steel.py
python update_itom_steel_input.py   # ~35 s
```

Step 1 must run before step 2. The two scripts resolve their paths relative to their own location. Region names are set by the constants `SOURCE_REGION` and `MODEL_REGION` at the top of `demand_and_scrap_availability_itom_steel.py`.

See [`ITOM_demand_interface/CLAUDE.md`](ITOM_demand_interface/CLAUDE.md) for the detailed logic, comparison against the previous data and open points.

## 2. Energy prices and biomass availability (OPEN-PROM)

[`ITOM_Energy_Interface/`](ITOM_Energy_Interface) reads an OPEN-PROM Excel export and updates two things in the ITOM workbook:

1. **Energy prices** (`update_itom_steel_energy_prices.py`)
   - Prices are taken from the `PricesIndustry` sheet (variables `Price|Final Energy|Industry|Iron and Steel|*`) for the EU27 + UK countries.
   - They are converted from k$2015/toe to €/GJ and turned into price indices relative to the base year 2020.
   - The indices are applied by carrier variant:
     - `pipeline` (natural gas, electricity): transport cost by mode on the pipeline routes.
     - `pipeline_direct` (hydrogen): direct conversion, with the old transport modes removed.
     - `commodity` (coal, biomass): global commodity import technologies.
2. **Biomass limit** (`update_itom_steel_biomass_limit.py`)
   - Final energy use of biomass and waste in iron and steel is summed over the EU27 + UK countries and over each 5-year period.
   - It is converted from Mtoe to PJ and written to `TotalTechnologyAnnualActivityUpperLimit` for `import_biomass` in region `energy_center`.

Run the whole pipeline (prices first, then biomass, on one output file):

```bash
cd ITOM_Energy_Interface
python data_update_open_prom_to_itom_steel.py
```

Everything is controlled by [`config_itom_steel.yaml`](ITOM_Energy_Interface/config_itom_steel.yaml): input and output file names, years, conversion factors, the ISO-3 → ISO-2 country mapping, the carrier mapping and the biomass settings. With `export_intermediate_csv: true`, the price indices are also written to `itom_steel_open_prom_relative_changes.csv`.

## 3. Carbon price / ETS (emissions penalty)

[`ITOM_ETS_Interface/`](ITOM_ETS_Interface) contains `TRANSIENCE_ITOM_ets_interface.py`. It reads the EU CO₂ price projection from `CarbonPrices.xlsx` (sheet `CarbonPrices`, region `EU`) and updates the `EmissionsPenalty` sheet of the ITOM workbook.

The input is averaged over 5-year windows (2025 → 2023–2027, …, 2050 → 2048–2050). The penalty for 2030–2050 is then scaled from the existing 2025 value and rounded to whole numbers. The method is set by `CALCULATION_METHOD` at the top of the script:

- `DIRECT_SINGLE_YEAR` (default): `Value_t = Base_2025 × Mean_t / Input_2025`
- `CHAIN`: `Value_t = Value_{t-1} × Mean_t / Mean_{t-1}`

```bash
cd ITOM_ETS_Interface
python TRANSIENCE_ITOM_ets_interface.py
```

The file paths (`INPUT_PATH`, `TARGET_PATH`, `OUTPUT_PATH`) are constants at the top of the script.

## Repository layout

```
.
├── ITOM_demand_interface/          MFA → demand and scrap availability
│   ├── demand_and_scrap_availability_itom_steel.py
│   ├── update_itom_steel_input.py
│   ├── mfa_itom_steel_product_mapping.xlsx
│   ├── output_to_itom/             MFA output CSVs (input)
│   ├── _model_data/                generated itom_steel_Demand.csv, itom_steel_Scrap.csv
│   ├── 4_Demand and Scrap availability_october26-update.R   (reference only)
│   └── CLAUDE.md
├── ITOM_Energy_Interface/          OPEN-PROM → energy prices, biomass limit
│   ├── data_update_open_prom_to_itom_steel.py   (entry point)
│   ├── update_itom_steel_energy_prices.py
│   ├── update_itom_steel_biomass_limit.py
│   ├── config_itom_steel.yaml
│   └── OPEN-PROM-Baseline-*.xlsx
└── ITOM_ETS_Interface/             CarbonPrices → emissions penalty
    ├── TRANSIENCE_ITOM_ets_interface.py
    └── CarbonPrices.xlsx
```

## Notes

- Never overwrite `ITOM_Steel_TRANSIENCE_baseline0_input.xlsx`. All scripts write to new files.
- `openpyxl` drops non-cell metadata from workbooks, such as the Google-Sheets `customXml` parts. Cell contents, formatting and images are kept.
- Check the constants at the top of each script (paths, regions, units) before running them in a new folder layout or with new MIC3 output.

## Authors

Developed in the TRANSIENCE project. The demand and scrap logic builds on an R script by Süheyb Bilici and Alexander Kling.
