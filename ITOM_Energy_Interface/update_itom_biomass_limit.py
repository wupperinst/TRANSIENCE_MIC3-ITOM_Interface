import pandas as pd
import openpyxl
import yaml
from pathlib import Path

def load_config(config_path: str = "config.yaml") -> dict:
    """Load settings and parameters from external YAML configuration."""
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

def get_col_idx(headers: list, target_names: list[str]) -> int:
    """Returns the 1-based index of the first header matching target names."""
    headers_upper = [str(h or "").strip().upper() for h in headers]
    for name in target_names:
        if name.upper() in headers_upper:
            return headers_upper.index(name.upper()) + 1
    raise ValueError(f"Could not find any of {target_names} in sheet headers: {headers}")

def extract_and_aggregate_biomass_pj(config: dict) -> dict[int, float]:
    """Reads OPEN-PROM biomass Excel data, converts Mtoe to PJ, and aggregates EU27+UK."""
    bio_cfg = config["biomass_settings"]
    input_file = Path(config["paths"]["open_prom_excel"])
    
    df = pd.read_excel(input_file, sheet_name=bio_cfg["sheet_name"])
    
    target_var = bio_cfg["variable_name"]
    eu_regions = bio_cfg["european_regions"]
    
    filtered_df = df[(df["Variable"] == target_var) & (df["Region"].isin(eu_regions))].copy()
    
    if filtered_df.empty:
        raise ValueError(f"No matching biomass data found for '{target_var}' in EU regions.")

    mtoe_to_pj = bio_cfg["mtoe_to_pj_factor"]
    period_totals_pj = {}

    for target_year, period_years in bio_cfg["time_periods"].items():
        valid_cols = [col for col in period_years if col in filtered_df.columns]
        total_mtoe = filtered_df[valid_cols].sum().sum()
        total_pj = round(total_mtoe * mtoe_to_pj)
        period_totals_pj[int(target_year)] = total_pj

    return period_totals_pj

def run_biomass_limit_update(config: dict, aggregated_biomass_pj: dict[int, float] = None, source_excel_path: Path = None, target_excel_path: Path = None) -> None:
    """Updates TotalTechnologyAnnualActivityUpperLimit worksheet in ITOM Excel file with PJ values."""
    excel_in = source_excel_path or Path(config["paths"]["itom_input_excel"])
    excel_out = target_excel_path or Path(config["paths"]["itom_output_excel"])
    bio_cfg = config["biomass_settings"]
    
    if aggregated_biomass_pj is None:
        aggregated_biomass_pj = extract_and_aggregate_biomass_pj(config)
        
    target_tech = bio_cfg["technology_name"]
    target_region = bio_cfg["region_name"]
    
    wb = openpyxl.load_workbook(excel_in)
    ws_limit = wb["Total...AnnualActivityUpLimit"]
    
    headers = [cell.value for cell in ws_limit[1]]
    reg_col = get_col_idx(headers, ["REGION", "Region"])
    tech_col = get_col_idx(headers, ["TECHNOLOGY", "Technology"])
    yr_col = get_col_idx(headers, ["YEAR", "Year"])
    limit_col = get_col_idx(headers, ["TotalTechnologyAnnualActivityUpperLimit", "TOTALTECHNOLOGYANNUALACTIVITYUPPERLIMIT"])

    updated_years = set()

    for r in range(2, ws_limit.max_row + 1):
        reg_val = str(ws_limit.cell(row=r, column=reg_col).value or "").strip()
        tech_val = str(ws_limit.cell(row=r, column=tech_col).value or "").strip()
        yr_val = ws_limit.cell(row=r, column=yr_col).value
        
        if reg_val == target_region and tech_val == target_tech and yr_val is not None:
            yr = int(yr_val)
            if yr in aggregated_biomass_pj:
                ws_limit.cell(row=r, column=limit_col).value = aggregated_biomass_pj[yr]
                updated_years.add(yr)

    for yr, new_val in aggregated_biomass_pj.items():
        if yr not in updated_years:
            ws_limit.append({
                reg_col: target_region, tech_col: target_tech,
                yr_col: yr, limit_col: new_val
            })

    wb.save(excel_out)
    print(f"[Biomass] Updated ITOM Excel saved to: {excel_out.resolve()}")

if __name__ == "__main__":
    cfg = load_config("config.yaml")
    biomass_data_pj = extract_and_aggregate_biomass_pj(cfg)
    run_biomass_limit_update(cfg, biomass_data_pj)