import re
import pandas as pd
import openpyxl
import yaml
from pathlib import Path

def load_config(config_path: str = "config_itom_steel.yaml") -> dict:
    """Load settings and mapping definitions from external YAML file."""
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

def get_col_idx(headers: list, target_names: list[str]) -> int:
    """Returns the 1-based index of the first matching header (case-insensitive)."""
    headers_upper = [str(h or "").strip().upper() for h in headers]
    for name in target_names:
        if name.upper() in headers_upper:
            return headers_upper.index(name.upper()) + 1
    raise ValueError(f"Could not find any of {target_names} in sheet headers: {headers}")

def get_optional_col_idx(headers: list, target_names: list[str]) -> int | None:
    """Returns the 1-based index of a matching header, or None if not found."""
    headers_upper = [str(h or "").strip().upper() for h in headers]
    for name in target_names:
        if name.upper() in headers_upper:
            return headers_upper.index(name.upper()) + 1
    return None

def compute_open_prom_price_indices(config: dict) -> pd.DataFrame:
    """Extracts OPEN-PROM prices from Excel, converts units, and calculates cumulative indices."""
    input_file = Path(config["paths"]["open_prom_excel"])
    sheet_name = config["paths"].get("open_prom_price_sheet", "PricesIndustry")
    
    df = pd.read_excel(input_file, sheet_name=sheet_name)
    
    col_region = config["columns"]["region"]
    col_variable = config["columns"]["variable"]
    base_year = int(config["base_year"])
    target_years = [int(y) for y in config["years"]]
    
    carrier_map = config["carrier_mapping"]
    country_map = config["country_mapping"]
    conv_factor = float(config.get("conversion_settings", {}).get("k_usd2015_per_toe_to_eur_per_gj", 21.52))
    
    df = df[df[col_variable].isin(carrier_map.keys()) & df[col_region].isin(country_map.keys())].copy()
    df["itom_carrier"] = df[col_variable].map(carrier_map)
    df["country_code_2"] = df[col_region].apply(lambda x: country_map[x]["code2"])
    
    year_cols = [col for col in df.columns if str(col) in map(str, target_years) or col in target_years]
    id_vars = ["country_code_2", "itom_carrier"]
    
    df_long = pd.melt(df, id_vars=id_vars, value_vars=year_cols, var_name="YEAR", value_name="price_raw")
    df_long["YEAR"] = df_long["YEAR"].astype(int)
    
    # Calculate absolute converted price in €/GJ
    df_long["price_eur_per_gj"] = df_long["price_raw"] * conv_factor
    
    df_long = df_long.sort_values(by=["country_code_2", "itom_carrier", "YEAR"])
    df_long["rel_change_prev_step"] = df_long.groupby(["country_code_2", "itom_carrier"])["price_raw"].pct_change()

    base_prices = (
        df_long[df_long["YEAR"] == base_year]
        .rename(columns={"price_raw": "base_price"})[["country_code_2", "itom_carrier", "base_price"]]
    )
    
    df_indexed = pd.merge(df_long, base_prices, on=["country_code_2", "itom_carrier"], how="left")
    df_indexed["price_index"] = df_indexed["price_raw"] / df_indexed["base_price"]
    
    return df_indexed

def run_energy_prices_update(config: dict, open_prom_indices: pd.DataFrame = None, source_excel_path: Path = None, target_excel_path: Path = None) -> None:
    """Updates the ITOM steel Excel file for energy prices (Pipeline Index, Commodity Index, and Pipeline Direct)."""
    excel_in = source_excel_path or Path(config["paths"]["itom_input_excel"])
    excel_out = target_excel_path or Path(config["paths"]["itom_output_excel"])
    base_year = int(config["base_year"])
    target_years = [int(y) for y in config["years"]]
    
    if open_prom_indices is None:
        open_prom_indices = compute_open_prom_price_indices(config)
    
    wb = openpyxl.load_workbook(excel_in)
    audit_log = []

    # Identify carriers for each variant
    pipeline_carriers_index = {
        details["itom_product_name"]: carrier
        for carrier, details in config["itom_carrier_settings"].items()
        if details.get("variant") == "pipeline"
    }
    
    pipeline_carriers_direct = {
        details["itom_product_name"]: carrier
        for carrier, details in config["itom_carrier_settings"].items()
        if details.get("variant") == "pipeline_direct"
    }

    ws_route = wb["TransportRoute_pipeline"]
    ws_cost = wb["TransportCostByMode"]
    ws_tm = wb["TRANSPORTMODE"]
    ws_cap = wb["TransportCapacity_pipeline"]
    ws_c2a = wb["TransportCapacityToActivity"]
    
    route_headers = [cell.value for cell in ws_route[1]]
    loc_indices = [i + 1 for i, h in enumerate(route_headers) if str(h or "").strip().upper() == "LOCATION"]
    loc1_col = loc_indices[0]
    loc2_col = loc_indices[1] if len(loc_indices) >= 2 else loc_indices[0]
    prod_col = get_col_idx(route_headers, ["product", "PRODUCT"])
    mode_col = get_col_idx(route_headers, ["transportmode", "TRANSPORTMODE"])
    r_year_col = get_col_idx(route_headers, ["YEAR", "Year"])

    cost_headers = [cell.value for cell in ws_cost[1]]
    c_mode_col = get_col_idx(cost_headers, ["transportmode", "TRANSPORTMODE"])
    c_year_col = get_col_idx(cost_headers, ["YEAR", "Year"])
    c_cost_col = get_col_idx(cost_headers, ["TransportCostByMode", "TRANSPORTCOSTBYMODE"])
    c_region_col = get_optional_col_idx(cost_headers, ["Region", "REGION"])

    tm_headers = [cell.value for cell in ws_tm[1]]
    tm_col = get_col_idx(tm_headers, ["TRANSPORTMODE", "transportmode"])

    cap_headers = [cell.value for cell in ws_cap[1]]
    cap_loc_indices = [i + 1 for i, h in enumerate(cap_headers) if str(h or "").strip().upper() == "LOCATION"]
    cap_loc1_col = cap_loc_indices[0]
    cap_loc2_col = cap_loc_indices[1] if len(cap_loc_indices) >= 2 else cap_loc_indices[0]
    cap_prod_col = get_col_idx(cap_headers, ["PRODUCT", "product"])
    cap_mode_col = get_col_idx(cap_headers, ["TRANSPORTMODE", "transportmode"])
    cap_year_col = get_col_idx(cap_headers, ["YEAR", "Year"])
    cap_val_col = get_col_idx(cap_headers, ["TransportCapacity", "TRANSPORTCAPACITY"])

    c2a_headers = [cell.value for cell in ws_c2a[1]]
    c2a_mode_col = get_col_idx(c2a_headers, ["TRANSPORTMODE", "transportmode"])
    c2a_val_col = get_col_idx(c2a_headers, ["TransportCapacityToActivity", "TRANSPORTCAPACITYTOACTIVITY"])

    # --- PURGE OLD MODES FOR DIRECT PIPELINE CARRIERS ---
    direct_product_names = set(pipeline_carriers_direct.keys())
    
    # 1. Purge from TransportCostByMode
    rows_to_delete = []
    for r in range(2, ws_cost.max_row + 1):
        m = str(ws_cost.cell(row=r, column=c_mode_col).value or "").lower()
        for prod_name in direct_product_names:
            if m.startswith(prod_name.lower()):
                rows_to_delete.append(r)
                break
    for r in reversed(rows_to_delete):
        ws_cost.delete_rows(r)

    # 2. Purge from TRANSPORTMODE
    rows_to_delete_tm = []
    for r in range(2, ws_tm.max_row + 1):
        m = str(ws_tm.cell(row=r, column=tm_col).value or "").lower()
        for prod_name in direct_product_names:
            if m.startswith(prod_name.lower()):
                rows_to_delete_tm.append(r)
                break
    for r in reversed(rows_to_delete_tm):
        ws_tm.delete_rows(r)

    # 3. Purge from TransportCapacityToActivity
    rows_to_delete_c2a = []
    for r in range(2, ws_c2a.max_row + 1):
        m = str(ws_c2a.cell(row=r, column=c2a_mode_col).value or "").lower()
        for prod_name in direct_product_names:
            if m.startswith(prod_name.lower()):
                rows_to_delete_c2a.append(r)
                break
    for r in reversed(rows_to_delete_c2a):
        ws_c2a.delete_rows(r)

    # Re-read existing costs & unique regions
    existing_costs = {}
    unique_regions = set()
    for r in range(2, ws_cost.max_row + 1):
        m = ws_cost.cell(row=r, column=c_mode_col).value
        y = ws_cost.cell(row=r, column=c_year_col).value
        v = ws_cost.cell(row=r, column=c_cost_col).value
        if c_region_col:
            reg = ws_cost.cell(row=r, column=c_region_col).value
            if reg:
                unique_regions.add(reg)
        if m is not None and y is not None:
            existing_costs[(str(m), int(y))] = v

    if not unique_regions:
        unique_regions = {None}

    all_pipeline_carriers = {**pipeline_carriers_index, **pipeline_carriers_direct}

    price_to_mode = {}
    max_mode_num = {}
    mode_prefixes = {}

    for product_name in all_pipeline_carriers.keys():
        prefix = f"{product_name.lower()}_"
        max_num = 0
        price_map = {}
        
        for (m_name, yr), cost_val in existing_costs.items():
            if m_name.lower().startswith(product_name.lower()):
                match = re.match(r"^([a-zA-Z]+_?)(\d+)$", m_name)
                if match:
                    prefix = match.group(1)
                    num = int(match.group(2))
                    if num > max_num:
                        max_num = num
                
                if cost_val is not None:
                    rounded_cost = int(round(cost_val))
                    price_map[rounded_cost] = m_name
                    
        price_to_mode[product_name] = price_map
        max_mode_num[product_name] = max_num
        mode_prefixes[product_name] = prefix

    base_prices = {}
    for r in range(2, ws_route.max_row + 1):
        prod_val = ws_route.cell(row=r, column=prod_col).value
        yr_val = ws_route.cell(row=r, column=r_year_col).value
        loc2_val = str(ws_route.cell(row=r, column=loc2_col).value or "")
        old_mode = ws_route.cell(row=r, column=mode_col).value
        
        if prod_val in pipeline_carriers_index and yr_val is not None and int(yr_val) == base_year:
            country_code = loc2_val[:2]
            b_cost = existing_costs.get((str(old_mode), base_year), 0.0)
            base_prices[(country_code, prod_val)] = b_cost

    new_modes_to_add = {}
    capacity_rows_to_add = []

    # Process routes for ALL pipeline carriers (Index & Direct)
    for r in range(2, ws_route.max_row + 1):
        prod_val = ws_route.cell(row=r, column=prod_col).value
        yr_val = ws_route.cell(row=r, column=r_year_col).value
        loc1_val = str(ws_route.cell(row=r, column=loc1_col).value or "")
        loc2_val = str(ws_route.cell(row=r, column=loc2_col).value or "")
        old_mode = ws_route.cell(row=r, column=mode_col).value
        
        if prod_val in all_pipeline_carriers and yr_val is not None:
            country_code = loc2_val[:2]
            carrier = all_pipeline_carriers[prod_val]
            yr = int(yr_val)
            
            country_indices = open_prom_indices[
                (open_prom_indices["country_code_2"] == country_code) &
                (open_prom_indices["itom_carrier"] == carrier) &
                (open_prom_indices["YEAR"] == yr)
            ]
            
            if not country_indices.empty:
                if prod_val in pipeline_carriers_direct:
                    abs_eur_gj = country_indices["price_eur_per_gj"].values[0]
                    new_price = int(round(abs_eur_gj))
                else:
                    base_p = base_prices.get((country_code, prod_val))
                    if base_p is None:
                        base_p = existing_costs.get((str(old_mode), base_year), 0.0)
                    idx = country_indices["price_index"].values[0]
                    new_price = int(round(base_p * idx))
                
                if new_price in price_to_mode[prod_val]:
                    assigned_mode = price_to_mode[prod_val][new_price]
                else:
                    max_mode_num[prod_val] += 1
                    next_num = max_mode_num[prod_val]
                    prefix = mode_prefixes[prod_val]
                    assigned_mode = f"{prefix}{next_num}"
                    
                    price_to_mode[prod_val][new_price] = assigned_mode
                    new_modes_to_add[assigned_mode] = (prod_val, new_price)
                
                ws_route.cell(row=r, column=mode_col).value = assigned_mode
                
                capacity_rows_to_add.append({
                    "loc1": loc1_val, "loc2": loc2_val,
                    "product": prod_val, "mode": assigned_mode, "year": yr
                })
                
                audit_log.append({
                    "Variant": "Pipeline_Direct" if prod_val in pipeline_carriers_direct else "Pipeline_Index",
                    "Carrier": carrier, "Country": country_code, "Year": yr,
                    "Old_Mode": old_mode, "Assigned_Mode": assigned_mode, "New_Price": new_price
                })

    for mode_name, (prod_val, price_val) in new_modes_to_add.items():
        for yr in target_years:
            for reg in unique_regions:
                row_data = {c_mode_col: mode_name, c_year_col: yr, c_cost_col: price_val}
                if c_region_col and reg is not None:
                    row_data[c_region_col] = reg
                ws_cost.append(row_data)

    existing_tm_set = {
        str(ws_tm.cell(row=r, column=tm_col).value)
        for r in range(2, ws_tm.max_row + 1)
        if ws_tm.cell(row=r, column=tm_col).value is not None
    }
    
    for mode_name in new_modes_to_add.keys():
        if mode_name not in existing_tm_set:
            ws_tm.append({tm_col: mode_name})
            existing_tm_set.add(mode_name)

    # --- UPDATE / APPEND TO TransportCapacityToActivity (Ratio = 5) ---
    existing_c2a_modes = {}
    for r in range(2, ws_c2a.max_row + 1):
        m_val = str(ws_c2a.cell(row=r, column=c2a_mode_col).value or "").strip()
        if m_val:
            existing_c2a_modes[m_val] = r

    for mode_name in new_modes_to_add.keys():
        if mode_name in existing_c2a_modes:
            target_r = existing_c2a_modes[mode_name]
            ws_c2a.cell(row=target_r, column=c2a_val_col).value = 5
        else:
            ws_c2a.append({c2a_mode_col: mode_name, c2a_val_col: 5})
            existing_c2a_modes[mode_name] = ws_c2a.max_row

    existing_cap_entries = {}
    for r in range(2, ws_cap.max_row + 1):
        l1 = str(ws_cap.cell(row=r, column=cap_loc1_col).value or "")
        l2 = str(ws_cap.cell(row=r, column=cap_loc2_col).value or "")
        p = str(ws_cap.cell(row=r, column=cap_prod_col).value or "")
        m = str(ws_cap.cell(row=r, column=cap_mode_col).value or "")
        y = ws_cap.cell(row=r, column=cap_year_col).value
        if y is not None:
            existing_cap_entries[(l1, l2, p, m, int(y))] = r

    for entry in capacity_rows_to_add:
        key = (entry["loc1"], entry["loc2"], entry["product"], entry["mode"], entry["year"])
        if key in existing_cap_entries:
            target_r = existing_cap_entries[key]
            ws_cap.cell(row=target_r, column=cap_val_col).value = 999
        else:
            ws_cap.append({
                cap_loc1_col: entry["loc1"], cap_loc2_col: entry["loc2"],
                cap_prod_col: entry["product"], cap_mode_col: entry["mode"],
                cap_year_col: entry["year"], cap_val_col: 999
            })
            existing_cap_entries[key] = ws_cap.max_row

    # ---------------------------------------------------------
    # VARIANT 2: Commodity Import Carrier Variable Costs
    # ---------------------------------------------------------
    ws_var = wb["VariableCost"]
    var_headers = [cell.value for cell in ws_var[1]]
    tech_col = get_col_idx(var_headers, ["TECHNOLOGY", "Technology"])
    v_year_col = get_col_idx(var_headers, ["YEAR", "Year"])
    v_cost_col = get_col_idx(var_headers, ["VariableCost", "VARIABLECOST"])
    
    commodity_carriers = {}
    for carrier, details in config["itom_carrier_settings"].items():
        if details.get("variant") == "commodity":
            t_names = details.get("technology_names") or details.get("technology_name")
            if isinstance(t_names, str):
                t_names = [t_names]
            if isinstance(t_names, list):
                for t_name in t_names:
                    commodity_carriers[str(t_name).strip()] = carrier

    avg_indices = open_prom_indices.groupby(["itom_carrier", "YEAR"])["price_index"].mean().reset_index()

    commodity_base_costs = {}
    for r in range(2, ws_var.max_row + 1):
        tech_val = str(ws_var.cell(row=r, column=tech_col).value or "").strip()
        yr_cell = ws_var.cell(row=r, column=v_year_col).value
        cost_val = ws_var.cell(row=r, column=v_cost_col).value
        
        if tech_val in commodity_carriers and yr_cell is not None:
            try:
                yr_val = int(yr_cell)
            except (ValueError, TypeError):
                continue
                
            if yr_val == base_year and cost_val is not None:
                commodity_base_costs[tech_val] = float(cost_val)

    modified_commodities_count = 0
    for r in range(2, ws_var.max_row + 1):
        tech_val = str(ws_var.cell(row=r, column=tech_col).value or "").strip()
        yr_cell = ws_var.cell(row=r, column=v_year_col).value
        
        if tech_val in commodity_carriers and yr_cell is not None:
            try:
                yr_val = int(yr_cell)
            except (ValueError, TypeError):
                continue

            carrier = commodity_carriers[tech_val]
            base_c = commodity_base_costs.get(tech_val)
            if base_c is None:
                raw_cell_cost = ws_var.cell(row=r, column=v_cost_col).value
                base_c = float(raw_cell_cost) if raw_cell_cost is not None else 0.0

            idx_row = avg_indices[(avg_indices["itom_carrier"] == carrier) & (avg_indices["YEAR"] == yr_val)]
            
            if not idx_row.empty:
                idx = float(idx_row["price_index"].values[0])
                
                if base_c <= 5.0:
                    new_cost = round(base_c * idx, 4)
                else:
                    new_cost = int(round(base_c * idx))
                
                ws_var.cell(row=r, column=v_cost_col).value = new_cost
                modified_commodities_count += 1
                
                audit_log.append({
                    "Variant": "Commodity", "Carrier": carrier, "Country": "ALL (Avg)",
                    "Year": yr_val, "Old_Mode": tech_val, "Assigned_Mode": tech_val, "New_Price": new_cost
                })

    wb.save(excel_out)
    print(f"[Prices] Updated ITOM steel Excel saved to: {excel_out.resolve()} (Modified Commodity Entries: {modified_commodities_count})")

if __name__ == "__main__":
    cfg = load_config("config_itom_steel.yaml")
    indices_df = compute_open_prom_price_indices(cfg)
    if cfg.get("export_intermediate_csv", False):
        csv_out_path = Path(cfg["paths"]["output_relative_changes"])
        indices_df.to_csv(csv_out_path, index=False)
    run_energy_prices_update(cfg, indices_df)