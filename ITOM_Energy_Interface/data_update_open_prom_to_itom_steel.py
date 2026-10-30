from pathlib import Path
import yaml

# Import functions from consistently named modules
from update_itom_steel_energy_prices import compute_open_prom_price_indices, run_energy_prices_update
from update_itom_steel_biomass_limit import extract_and_aggregate_biomass_pj, run_biomass_limit_update

def load_config(config_path: str = "config_itom_steel.yaml") -> dict:
    """Load settings and parameters from external YAML configuration."""
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

def run_pipeline():
    """Executes both Energy Price and Biomass updates sequentially on a single ITOM steel Excel output file."""
    print("="*60)
    print("   STARTING INTEGRATED OPEN-PROM TO ITOM STEEL DATA PIPELINE")
    print("="*60)
    
    cfg = load_config("config_itom_steel.yaml")
    
    itom_input = Path(cfg["paths"]["itom_input_excel"])
    itom_output = Path(cfg["paths"]["itom_output_excel"])

    # ---------------------------------------------------------
    # STEP 1: Energy Prices Update
    # ---------------------------------------------------------
    print("\n>>> Step 1/2: Processing Energy Prices...")
    price_indices_df = compute_open_prom_price_indices(cfg)
    
    if cfg.get("export_intermediate_csv", False):
        csv_out_path = Path(cfg["paths"]["output_relative_changes"])
        price_indices_df.to_csv(csv_out_path, index=False)
        print(f"    Exported intermediate price indices to: {csv_out_path.resolve()}")
        
    run_energy_prices_update(
        config=cfg, 
        open_prom_indices=price_indices_df, 
        source_excel_path=itom_input, 
        target_excel_path=itom_output
    )

    # ---------------------------------------------------------
    # STEP 2: Biomass Availability Update (Values in PJ)
    # ---------------------------------------------------------
    print("\n>>> Step 2/2: Processing Biomass Availability (in PJ)...")
    biomass_data_pj = extract_and_aggregate_biomass_pj(cfg)
    
    run_biomass_limit_update(
        config=cfg, 
        aggregated_biomass_pj=biomass_data_pj, 
        source_excel_path=itom_output, 
        target_excel_path=itom_output
    )

    print("\n" + "="*60)
    print("   SUCCESSFULLY COMPLETED ALL UPDATES!")
    print(f"   Final Combined ITOM Steel File: {itom_output.resolve()}")
    print("="*60)

if __name__ == "__main__":
    run_pipeline()