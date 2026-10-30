import pandas as pd
import openpyxl

# ==============================================================================
# CONFIGURATION
# ==============================================================================
# Select calculation method for growth rates:
#  - "CHAIN"              : Step-by-step chain index using period means
#                           Value_t = Value_{t-1} * (Mean_t / Mean_{t-1})
#  - "DIRECT_SINGLE_YEAR" : Direct ratio relative to single 2025 input value
#                           Value_t = BaseValue_2025 * (Mean_t / Input_2025)
CALCULATION_METHOD = "DIRECT_SINGLE_YEAR"

INPUT_PATH = "CarbonPrices.xlsx"
TARGET_PATH = "ITOM_Steel_TRANSIENCE_baseline0_input_MFA-Energy-update.xlsx"
OUTPUT_PATH = "ITOM_Steel_TRANSIENCE_baseline0_input_MFA-Energy-ETS_update.xlsx"
SHEET_NAME = "EmissionsPenalty"


def update_emissions_penalty(
    input_path: str = INPUT_PATH,
    target_path: str = TARGET_PATH,
    output_path: str = OUTPUT_PATH,
    sheet_name: str = SHEET_NAME,
    method: str = CALCULATION_METHOD
) -> None:
    """
    Reads CO2 price projections, calculates growth rates based on the chosen method,
    rounds updated penalty values to whole numbers (integers), and saves the target sheet.
    """
    # 1. Load input dataset and filter for EU region
    df_input = pd.read_excel(input_path, sheet_name="CarbonPrices")
    df_eu = df_input[df_input["Region"] == "EU"]
    
    if df_eu.empty:
        raise ValueError("No rows matching Region='EU' were found in the input file.")
    
    input_series = df_eu.iloc[0]
    
    # 2. Extract single year 2025 input value (needed for direct calculation method)
    val_2025_col = [c for c in input_series.index if str(c) in ["2025", 2025]]
    if not val_2025_col:
        raise KeyError("Column for year 2025 missing from input file.")
    input_val_2025 = float(input_series[val_2025_col[0]])

    # 3. Define 5-year rolling windows for input period mean calculation
    window_mapping = {
        2025: range(2023, 2028),
        2030: range(2028, 2033),
        2035: range(2033, 2038),
        2040: range(2038, 2043),
        2045: range(2043, 2048),
        2050: range(2048, 2051)  # Truncated 3-year period (ends in 2050)
    }
    
    input_means = {}
    for year, yr_range in window_mapping.items():
        valid_cols = [
            col for col in input_series.index 
            if (isinstance(col, int) and col in yr_range) or 
               (isinstance(col, str) and col.isdigit() and int(col) in yr_range)
        ]
        if not valid_cols:
            raise KeyError(f"No valid year columns found in input file for period range {yr_range}.")
            
        input_means[year] = float(input_series[valid_cols].astype(float).mean())

    # 4. Load target workbook preserving all existing sheets and formatting
    wb = openpyxl.load_workbook(target_path)
    
    if sheet_name not in wb.sheetnames:
        raise KeyError(f"Sheet '{sheet_name}' was not found in target workbook.")
        
    ws = wb[sheet_name]
    
    header = [cell.value for cell in ws[1]]
    try:
        year_col_idx = header.index("YEAR") + 1
        penalty_col_idx = header.index("EmissionsPenalty") + 1
    except ValueError as e:
        raise KeyError(f"Required columns 'YEAR' or 'EmissionsPenalty' missing from '{sheet_name}'.") from e
    
    # 5. Extract 2025 baseline value from target sheet
    base_val_2025 = None
    for row in range(2, ws.max_row + 1):
        cell_year = ws.cell(row=row, column=year_col_idx).value
        if cell_year == 2025 or cell_year == "2025":
            base_val_2025 = float(ws.cell(row=row, column=penalty_col_idx).value)
            break
            
    if base_val_2025 is None:
        raise ValueError("Baseline value for year 2025 was not found in target sheet.")

    # 6. Compute rounded penalty values based on selected calculation method
    target_years = [2030, 2035, 2040, 2045, 2050]
    updated_values = {2025: int(round(base_val_2025))}
    
    method_key = method.upper().strip()
    
    if method_key == "CHAIN":
        # Chain index option: relative step change between consecutive period means
        prev_year = 2025
        for yr in target_years:
            change_rate = input_means[yr] / input_means[prev_year]
            raw_val = updated_values[prev_year] * change_rate
            updated_values[yr] = int(round(raw_val))
            prev_year = yr

    elif method_key == "DIRECT_SINGLE_YEAR":
        # Direct option: ratio of period mean relative to single-year 2025 input value
        for yr in target_years:
            change_rate = input_means[yr] / input_val_2025
            raw_val = base_val_2025 * change_rate
            updated_values[yr] = int(round(raw_val))

    else:
        raise ValueError(f"Invalid method '{method}'. Choose 'CHAIN' or 'DIRECT_SINGLE_YEAR'.")

    # 7. Update target worksheet cells with integer values
    for row in range(2, ws.max_row + 1):
        cell_year = ws.cell(row=row, column=year_col_idx).value
        try:
            year_val = int(cell_year)
        except (ValueError, TypeError):
            continue
            
        if year_val in target_years:
            ws.cell(row=row, column=penalty_col_idx).value = updated_values[year_val]

    # 8. Save updated workbook
    wb.save(output_path)
    print(f"Process completed using '{method_key}' method. Saved as: {output_path}")

if __name__ == "__main__":
    update_emissions_penalty()