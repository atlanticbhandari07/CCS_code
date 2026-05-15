import os 
import time 
import math 
import csv 
from datetime import datetime 
import win32com.client as win32 
import pythoncom
import matplotlib.pyplot as plt
import pandas as pd
import numpy as np

# Case file path from ASPEN HYSYS v14 
HYSYS_CASE_PATH = r"d:\BUUU\New simulation with new CAPEX sheet base case with spreadsheet 15 DelTmin.hsc" 
# Names of Units and Streams in Case File simulation  
HX_NAME = "E-100"                               # Main Lean/Rich HEX 
AMINE_COOLER_NAME = "E-101"                     # Amine Cooler
PUMP_100_NAME = "P-100"                         # Pump 100
PUMP_101_NAME = "P-101"                         # Pump 101

CO2_PRODUCT_STREAM_NAME = "CO2 product"         # CO2 stream Name  
REBOILER_ENERGY_STREAM_NAME = "Reboiler duty"   # Reboiler Duty Stream 
CONDENSER_ENERGY_STREAM_NAME = "Condenser duty" # Condenser Duty Stream 

# Adjust block Function control DelTmin 
DELTMIN_ADJUST_NAME = "ADJ-2" 
DELTMIN_VALUES = list (range(5, 16, 1)) 

# Name of Streams for E-100 Heat Exchanger (Lean/Rich)
LEAN_HOT_IN_NAME = "Lean MEA 2"
LEAN_HOT_OUT_NAME = "Lean MEA 3" 
RICH_COLD_IN_NAME = "Rich MEA 2" 
RICH_COLD_OUT_NAME = "Rich MEA 3" 

# Convergence  
MAX_WAIT_S = 120 
POLL_S = 1.0 
EXTRA_SETTLE_S = 1.0 

# Output 
OUT_CSV = f"co2_capture_cost_optimization_overall_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv" 
U_KW_PER_M2K = 1.5 
C1_SS_2021_EUR  = 1662000.0   # Referece  base cost for ss
A1_M2=            2747.0       # Reference area (m2) for base cost
EXPONENT_E      = 1.0          # Scaling exponent (1.0 = linear with area) 
SS_TO_CS_FACTOR = 1.75         # material factor
INSTALL_FACTOR  = 5.3          # E-100 installation factor 
CEPCI_2026      = 830.0        #CEPCI for 2026
CEPCI_2021      = 820.97       # CEPCI for 2021

OPERATING_HOURS_PER_YEAR = 8000.0 
STEAM_PRICE_EUR_PER_KWH = 0.015 
ELECTRICITY_PRICE_EUR_PER_KWH = 0.05 
ANNUALIZATION_FACTOR_AF = 11.15 
MAINTENANCE_FRACTION = 0.05 
MEA_SOLVENT_COST_EUR_PER_YEAR = 708571.0 

# Net Present Value (NPV) parameters 
NPV_PROJECT_YEARS = 25           # Project life in years 
NPV_DISCOUNT_RATE = 0.075        # Discount rate (7.5%) 
# Present-Value Annuity Factor = (1 - (1+r)^-n) / r 
NPV_PV_ANNUITY_FACTOR = (1 - (1 + NPV_DISCOUNT_RATE) ** (-NPV_PROJECT_YEARS)) /NPV_DISCOUNT_RATE

# Pump base cost and installation factors 
PUMP_REF_CAPACITY_M3PH = 940.0 
PUMP_REF_COST_2021_EUR = 42307.7 
PUMP_COST_EXPONENT = 0.6 
P100_INSTALL_FACTOR = 7.580 
P101_INSTALL_FACTOR = 8.90 

# Base cost, Installation factor and LMTD for Condenser, Reboiler and Amine Cooler
HEX_AREA_REF_COST_2021_EUR = C1_SS_2021_EUR / SS_TO_CS_FACTOR 
CONDENSER_U_KW_M2K = 1.0 
CONDENSER_LMTD_K = 71.14 
CONDENSER_INSTALL_FACTOR = 9.480 
REBOILER_U_KW_M2K = 1.2 
REBOILER_LMTD_K = 27.18 
REBOILER_INSTALL_FACTOR = 5.30 
COOLER_U_KW_M2K = 0.8 
COOLER_LMTD_K = 27.0 
COOLER_INSTALL_FACTOR = 5.30 

# Base cost and installation factors for Absorber and Desorber 
ABSORBER_PURCHASE_COST_2021_EUR = 2994285.0      # Equipment cost for Absorber
ABSORBER_INSTALL_FACTOR = 3.620 
DESORBER_PURCHASE_COST_2021_EUR = 4385258.0      # Equipment cost for Desorber
DESORBER_INSTALL_FACTOR = 3.330  

# Defining functions 
def safe_getattr(obj, attr_name):
    """Return obj.attr_name if exists else raise a clear error."""
    if not hasattr(obj, attr_name):
        raise AttributeError(
            f"Object {obj} has no attribute '{attr_name}'. Check HYSYS COM API & names."
        )
    return getattr(obj, attr_name)

def wait_for_convergence(case, max_wait_s=300, poll_s=2): 
    """ 
    Waits for HYSYS solver to finish. 
    Uses case.Solver.IsSolving as the primary flag. 
    """ 
    start_time = time.time() 
    print(" Waiting for HYSYS solver...", end="", flush=True) 
 
    solver = None 
    try: 
        solver = case.Solver 
    except Exception: 
        pass 
 
    time.sleep(1) 
    while time.time() - start_time < max_wait_s: 
        is_solving = False 
        try: 
            if solver: 
                is_solving = solver.IsSolving 
            else: 
                is_solving = case.Solver.IsSolving 
        except Exception: 
            is_solving = True 
 
        if not is_solving: 
            print(" Solved.") 
            return True 
 
        print(".", end="", flush=True) 
        time.sleep(poll_s) 
 
    print(" Timeout.") 
    return False 
 
def get_component_massflow_kgph(stream, comp_name="CO2"): 
    """ 
    Read component mass flow (kg/h) for comp_name from a HYSYS Material Stream. 
    Dynamically finds the component index and uses ComponentMassFlowValue array. 
    """ 
    try: 
        fp = stream.FluidPackage 
        comps = fp.Components 
        comp_idx = -1 
        for i in range(comps.Count): 
            if comps.Item(i).name.upper() == comp_name.upper(): 
                comp_idx = i 
                break 
 
        if comp_idx == -1: 
            raise RuntimeError(f"Component '{comp_name}' not found in Fluid Package for stream '{stream.name}'.") 
 
        if hasattr(stream, "ComponentMassFlowValue"): 
            cmf_tuple = stream.ComponentMassFlowValue 
            if len(cmf_tuple) > comp_idx: 
                internal_rate = float(cmf_tuple[comp_idx]) 
                return internal_rate * 3600.0 
 
        total_kgph = 0.0 
        mf = stream.MassFlow 
        if hasattr(mf, "GetValue"): 
            total_kgph = float(mf.GetValue("kg/h")) 
        else: 
            total_kgph = float(mf.Value if hasattr(mf, "Value") else mf) * 3600.0 
 
        co2_frac = 0.0 
        if hasattr(stream, "ComponentMassFractionValue"): 
            cmfv_tuple = stream.ComponentMassFractionValue 
            if len(cmfv_tuple) > comp_idx: 
                co2_frac = float(cmfv_tuple[comp_idx]) 
 
        if total_kgph > 0 and co2_frac > 0: 
            return total_kgph * co2_frac 
 
    except Exception as e: 
        print(f" Warning: Error in get_component_massflow_kgph for {comp_name}: {e}") 
        raise RuntimeError(f"Could not read component mass flow for '{comp_name}' from stream '{stream.name}'. " "Check HYSYS COM connectivity and component names.") 
 
def get_stream_massflow_kgph(stream): 
    """Returns CO2 component mass flow (kg/h) from the CO2 product stream.""" 
    return get_component_massflow_kgph(stream, "CO2") 
 
def get_energy_kjph(energy_stream): 
    """Reads energy stream heat flow in kJ/h.""" 
    candidates = ["HeatFlow", "Duty", "EnergyFlow", "Q"] 
    for c in candidates: 
        if hasattr(energy_stream, c): 
            v = getattr(energy_stream, c) 
            if hasattr(v, "GetValue"): 
                try: 
                    return float(v.GetValue("kJ/h")) 
                except Exception: 
                    pass 
            if hasattr(v, "Value"): 
                try: 
                    return float(v.Value) 
                except Exception: 
                    pass 
            try: 
                return float(v) 
            except Exception:
                pass
    raise RuntimeError("Could not read energy stream heat flow in kJ/h. Check stream property names.") 
def get_hx_duty_kjph(hx): 
    """Reads HX duty in kJ/h from unit op.""" 
    candidates = ["Duty", "HeatFlow", "Q", "Energy"] 
    for c in candidates: 
        if hasattr(hx, c): 
            v = getattr(hx, c) 
            try: 
                if hasattr(v, "GetValue"): 
                    return float(v.GetValue("kJ/h")) 
                if hasattr(v, "Value"): 
                    return float(v.Value) 
                return float(v) 
            except Exception: 
                continue 
    raise RuntimeError("Could not read HX duty in kJ/h. Check HX property names in HYSYS.") 
 
def get_hx_lmtd(hx): 
    """Reads LMTD directly from HYSYS HX object.""" 
    candidates = ["Lmtd", "LMTD", "LogMeanTempDiff", "LogMeanTemperatureDifference"] 
    for c in candidates: 
        if hasattr(hx, c): 
            v = getattr(hx, c) 
            try: 
                if hasattr(v, "GetValue"): 
                    return float(v.GetValue("C")) 
                if hasattr(v, "Value"): 
                    return float(v.Value) 
                return float(v) 
            except Exception: 
                continue 
    return float("nan") 
 
def _get_stream_temp_C(stream): 
    """Best-effort read temperature in degC from a material stream.""" 
    if hasattr(stream, "Temperature"): 
        T = stream.Temperature 
        if hasattr(T, "GetValue"): 
            return float(T.GetValue("C")) 
        if hasattr(T, "Value"): 
            return float(T.Value) 
        return float(T) 
    raise RuntimeError("Stream has no Temperature property.") 
 
def calc_area_m2(Q_Kw, U_Kw_m2k, LMTD_K): 
    """A = Q / (U * LMTD).""" 
    if U_Kw_m2k <= 0 or LMTD_K <= 0: 
        return float("nan") 
    return Q_Kw / (U_Kw_m2k * LMTD_K) 
 
# Defining functions for CAPEX calculation and updating for 2026 using CEPCI Index 
def calc_capex_2026_eur(area_m2): 
    """ 
    E-100 (lean/rich HX) installed CAPEX escalated to 2026. 
    Steps: 
    1. Scale purchased SS316 cost (2021 basis) with area 
    2. Convert SS316 -> CS (material factor) 
    3. Apply installation factor 
    4. Escalate 2021 -> 2026 via CEPCI 
    """ 
    if area_m2 <= 0: 
        return float("nan") 
    C2_SS_2021 = C1_SS_2021_EUR * ((area_m2 / A1_M2) ** EXPONENT_E) 
    C2_CS_2021 = C2_SS_2021 / SS_TO_CS_FACTOR 
    capex_2021 = C2_CS_2021 * INSTALL_FACTOR 
    capex_2026 = capex_2021 * (CEPCI_2026 / CEPCI_2021) 
    return capex_2026 
 
# Defining functions for Annual OPEX calculation 
def calc_annual_opex_eur_per_year(Q_reb_kjph): 
    """ 
    Original reboiler OPEX logic kept unchanged. 
    Annual OPEX = (Q_reb(kJ/h) / 3600 -> kW) * operating hours * steam price 
    """ 
    Q_reb_kw = Q_reb_kjph / 3600.0 
    annual_energy_kwh = Q_reb_kw * OPERATING_HOURS_PER_YEAR 
    return annual_energy_kwh * STEAM_PRICE_EUR_PER_KWH 
 
# Defining functions for CO2 calculation 
def calc_annual_co2_tpy(co2_kgph): 
    """kg/h * 8000 h/yr / 1000 = t/yr""" 
    return (co2_kgph * OPERATING_HOURS_PER_YEAR) / 1000.0 
 
# Defining functions for cost escalation 
def escalate_cost_2021_to_2026(cost_2021): 
    return cost_2021 * (CEPCI_2026 / CEPCI_2021) 
 
def calc_area_based_purchase_cost_2021(area_m2, 
ref_cost_2021=HEX_AREA_REF_COST_2021_EUR): 
    if area_m2 <= 0: 
        return float("nan") 
    return ref_cost_2021 * ((area_m2 / A1_M2) ** EXPONENT_E) 
 
def calc_total_installed_cost_2026_from_purchase_2021(purchase_cost_2021, 
install_factor): 
    if purchase_cost_2021 <= 0: 
        return float("nan") 
    return escalate_cost_2021_to_2026(purchase_cost_2021 * install_factor) 
 
def _try_get_value(prop, units): 
    if hasattr(prop, "GetValue"): 
        for unit in units: 
            try: 
                return float(prop.GetValue(unit)) 
            except Exception: 
                pass 
    if hasattr(prop, "Value"): 
        try: 
            return float(prop.Value) 
        except Exception: 
            pass 
    try: 
        return float(prop) 
    except Exception: 
        pass 
    raise RuntimeError(f"Could not read value for units={units}") 
 
def get_numeric_property(obj, prop_candidates, unit_candidates): 
    errors = [] 
    for name in prop_candidates: 
        if hasattr(obj, name): 
            prop = getattr(obj, name) 
            try: 
                return _try_get_value(prop, unit_candidates) 
            except Exception as e: 
                errors.append(f"{name}: {e}") 
    raise RuntimeError( 
        f"Could not read any of properties {prop_candidates} from object '{getattr(obj, 'Name', 
obj)}'. " 
        f"Tried units {unit_candidates}. Details: {errors}" 
    ) 
 
def get_pump_capacity_m3ph(pump): 
    """ 
    Read pump volumetric capacity in m3/h. 
    In HYSYS, Pump 100 exposes a 'Capacity' variable (shown as m3/hr in UI). 
    Tries the property directly on the pump first, then falls back to the 
    pump's connected feed (inlet) stream volumetric flow. 
    """ 
    # 1) Try reading Capacity (and synonyms) directly from the pump object 
    candidates = [ 
        "Capacity", "capacity", 
        "PumpCapacity", "RatedCapacity", "VolumetricFlow", "LiquidFlow", 
    ] 
    # HYSYS COM GetValue accepts "m3/h" internally; "m3/hr" is the display unit. 
    # Try both so we handle whatever the COM layer accepts. 
    units = ["m3/h", "m3/hr", "m^3/h", "m^3/hr"] 
    errors = [] 
    for name in candidates: 
        if hasattr(pump, name): 
            prop = getattr(pump, name) 
            try: 
                return _try_get_value(prop, units) 
            except Exception as e: 
                errors.append(f"{name}: {e}") 
 
    # 2) Fallback: read volumetric flow from the pump's feed (inlet) stream 
    feed_stream = None 
    for feed_attr in ("FeedStream", "Feed", "InletStream", "Inlet"): 
        if hasattr(pump, feed_attr): 
            try: 
                feed_stream = getattr(pump, feed_attr) 
                break 
            except Exception: 
                pass 
 
    if feed_stream is not None: 
        vol_candidates = [ 
            "ActualVolumeFlow", "VolumeFlow", "VolumetricFlow", 
            "ActualVolFlow", "VolFlow", 
        ] 
        for name in vol_candidates: 
            if hasattr(feed_stream, name): 
                prop = getattr(feed_stream, name) 
                try: 
                    return _try_get_value(prop, units) 
                except Exception as e: 
                    errors.append(f"feed.{name}: {e}") 
 
    raise RuntimeError( 
        f"Could not read pump capacity (m3/h) from pump '{getattr(pump, 'Name', pump)}'. " 
        f"Tried pump properties {candidates} and feed stream volumetric flow. " 
        f"Details: {errors}" 
    ) 
 
def get_pump_power_kw(pump): 
    """ 
    Read pump shaft power in kW. 
    In HYSYS, 'Total Power' is in the Performance tab of the pump. 
 
    HYSYS COM access strategy (in order): 
    1. Variables collection  -> pump.Variables.Item("Total Power").GetValue("kW") 
       This is the correct way to access any Performance-tab variable whose 
       name contains a space, since Python getattr/hasattr cannot reach them. 
    2. Energy stream connected to the pump (pump.EnergyFeed / EnergyStream) 
    3. Direct camelCase property names (TotalPower, Power, etc.) 
    """ 
    units_kw   = ["kW"] 
    units_kjph = ["kJ/h"] 
    errors = [] 
 
   # Reading variables from ASPEN HYSYS for pump 
    var_names = ["Total Power", "TotalPower", "Power", "Brake Power", "Actual Power"] 
    for vname in var_names: 
        for vcoll_attr in ("Variables", "variables"): 
            if hasattr(pump, vcoll_attr): 
                try: 
                    vcoll = getattr(pump, vcoll_attr) 
                    var = vcoll.Item(vname) 
                    try: 
                        return float(var.GetValue("kW")) 
                    except Exception: 
                        pass 
                    try: 
                        return float(var.GetValue("kJ/h")) / 3600.0 
                    except Exception: 
                        pass 
                    try: 
                        return float(var.Value) 
                    except Exception as e: 
                        errors.append(f"Variables['{vname}'].Value: {e}") 
                except Exception as e: 
                    errors.append(f"Variables['{vname}']: {e}") 
# Reading energy stream connected to the pump 
    energy_stream = None 
    for eattr in ("EnergyFeed", "EnergyStream", "EnergyConnection", 
                  "PowerFeed", "WorkStream"): 
        if hasattr(pump, eattr): 
            try: 
                energy_stream = getattr(pump, eattr) 
                break 
            except Exception: 
                pass 
 
    if energy_stream is not None: 
        for eprop in ("HeatFlow", "Power", "Duty", "EnergyFlow", "Q"): 
            if hasattr(energy_stream, eprop): 
                prop = getattr(energy_stream, eprop) 
                try: 
                    return _try_get_value(prop, units_kw) 
                except Exception as e: 
                    errors.append(f"energy_stream.{eprop}@kW: {e}") 
                try: 
                    return _try_get_value(prop, units_kjph) / 3600.0 
                except Exception as e: 
                    errors.append(f"energy_stream.{eprop}@kJ/h: {e}") 
 
    for name in ("TotalPower", "Power", "BrakePower", "ActualPower", "Energy", "Duty"): 
        if hasattr(pump, name): 
            prop = getattr(pump, name) 
            try: 
                return _try_get_value(prop, units_kw) 
            except Exception as e: 
                errors.append(f"{name}@kW: {e}") 
            try: 
                return _try_get_value(prop, units_kjph) / 3600.0 
            except Exception as e: 
                errors.append(f"{name}@kJ/h: {e}") 
 
    raise RuntimeError( 
        f"Could not read 'Total Power' (kW) from pump '{getattr(pump, 'Name', pump)}'. " 
        f"Tried Variables collection, energy stream, and direct properties. " 
        f"Details: {errors}" 
    ) 
 
def calc_pump_purchase_cost_2021(capacity_m3ph): 
    if capacity_m3ph <= 0: 
        return float("nan") 
    return PUMP_REF_COST_2021_EUR * ((capacity_m3ph / PUMP_REF_CAPACITY_M3PH) 
** PUMP_COST_EXPONENT) 
 
def calc_pump_annual_opex_eur_per_year(power_kw): 
    if power_kw < 0: 
        return float("nan") 
    return power_kw * OPERATING_HOURS_PER_YEAR *ELECTRICITY_PRICE_EUR_PER_KWH 
 
def calc_direct_column_capex_2026(purchase_cost_2021, install_factor): 
    return calc_total_installed_cost_2026_from_purchase_2021(purchase_cost_2021, 
install_factor) 
 
# Defining Plot functions 
def plot_results(results, prefix="results"): 
    """ 
    Generates four plots: 
    1. cost_vs_dtmin.png   (Capture Cost + E-100 Area) 
    2. capex_vs_dtmin.png  (Total Annualized CAPEX) 
    3. opex_vs_dtmin.png   (Total Annual OPEX) 
    4. npv_vs_dtmin.png    (NPV – always negative; less negative = better) 
    """ 
    if not results: 
        return 
 
    dtmin_vals = [r["DelTmin_C"] for r in results] 
    costs = [r["CO2_Captured_Cost_EUR_per_t"] for r in results] 
    areas = [r["HX_Area_m2"] for r in results] 
    capex = [r["AnnualizedCAPEX_EUR_per_y"] for r in results] 
    opex = [r["AnnualOPEX_EUR_per_y"] for r in results] 
    npv_vals = [r["NPV_EUR"] for r in results] 
    hoc_vals = [r["HOC_MJ_per_kg_CO2"] for r in results] 
 
    # ── Read manual capture cost data from Excel (Book2) ────────────
    manual_cost_file = r"C:\Users\AZ400_ATLANTIC\Desktop\Book2.xlsx"
    try:
        df_manual_cost = pd.read_excel(manual_cost_file)
        df_manual_cost["DelTmin(0C)"] = pd.to_numeric(df_manual_cost["DelTmin(0C)"], errors='coerce')
        df_manual_cost["CO2 CAPTURED COST (Euro/tons of CO2)"] = pd.to_numeric(
            df_manual_cost["CO2 CAPTURED COST (Euro/tons of CO2)"], errors='coerce')
        df_manual_cost.dropna(subset=["DelTmin(0C)", "CO2 CAPTURED COST (Euro/tons of CO2)"], inplace=True)
        manual_cost_dtmin = df_manual_cost["DelTmin(0C)"].tolist()
        manual_cost_vals  = df_manual_cost["CO2 CAPTURED COST (Euro/tons of CO2)"].tolist()
        has_manual_cost = True
        print(f"Manual cost data loaded from {manual_cost_file} ({len(manual_cost_dtmin)} points).")
    except Exception as e:
        print(f"WARNING: Could not load manual cost data from {manual_cost_file}: {e}")
        print("         Plotting Python results only for cost_vs_dtmin.")
        has_manual_cost = False
        manual_cost_dtmin, manual_cost_vals = [], []

    fig1, ax1 = plt.subplots(figsize=(11, 6))
    color_cost_py  = '#0000FF'   # bright blue – Python cost
    color_cost_man = '#FF0000'   # bright red  – Manual cost

    # Python capture cost (solid line)
    lns1 = ax1.plot(
        dtmin_vals, costs, marker='o', linestyle='-', color=color_cost_py,
        linewidth=2, label="CO₂ Capture Cost (Python)"
    )
    # Manual capture cost (dashed line)
    lns_man_cost = []
    if has_manual_cost:
        lns_mc = ax1.plot(
            manual_cost_dtmin, manual_cost_vals, marker='x', linestyle='--',
            color=color_cost_man, linewidth=2, label="CO₂ Capture Cost (Manual)"
        )
        lns_man_cost = lns_mc

    ax1.set_xlabel("ΔTmin (°C)", fontsize=14, color='black')
    ax1.set_ylabel("CO₂ Capture Cost (€/tCO₂)", fontsize=14, color='black')
    ax1.tick_params(axis='y', labelcolor='black', labelsize=14)
    ax1.tick_params(axis='x', labelcolor='black', labelsize=14)
    ax1.grid(True, linestyle='--', alpha=0.7)

    ax2 = ax1.twinx()
    color_area = '#0a6a04'       # dark green
    lns2 = ax2.plot(
        dtmin_vals, areas, marker='s', linestyle='--', color=color_area,
        linewidth=1.5, alpha=0.8, label="E-100 Area (m²)"
    )
    ax2.set_ylabel("E-100 Area (m²)", fontsize=14, color='black')
    ax2.tick_params(axis='y', labelcolor='black', labelsize=14)

    min_cost = min(costs)
    best_dtmin = dtmin_vals[costs.index(min_cost)]
    sc = ax1.scatter(
        best_dtmin, min_cost, color='#FF0000', s=200, zorder=6,
        edgecolors='black', linewidths=1.5,
        label=f"Optimal ΔTmin — Minimum Capture Cost ({int(best_dtmin)}°C)"
    )
    ax1.set_xticks(dtmin_vals)

    # Combined legend for both axes including optimal point
    lns_c = lns1 + lns_man_cost + lns2 + [sc]
    labs_c = [l.get_label() for l in lns_c]
    ax1.legend(lns_c, labs_c, loc='upper center', bbox_to_anchor=(0.5, -0.22),
               ncol=2, fontsize=11, framealpha=0.9, edgecolor='gray')
    fig1.tight_layout()
    fig1.subplots_adjust(bottom=0.25)
    plt.savefig("cost_vs_dtmin.png", dpi=150, bbox_inches='tight')
 
    # ── Stacked bar chart for CAPEX and OPEX (Total Annualized Cost) ──
    capex_m = np.array(capex) / 1e6
    opex_m = np.array(opex) / 1e6
    total_m = capex_m + opex_m

    fig_stack, ax_stack = plt.subplots(figsize=(11, 7))
    x_stack = np.arange(len(dtmin_vals))

    ax_stack.bar(x_stack, capex_m, color='#5DADE2', label='Annualized CAPEX', edgecolor='white', linewidth=0.6)
    ax_stack.bar(x_stack, opex_m, color='#E67E22', label='Annual OPEX', edgecolor='white', linewidth=0.6, bottom=capex_m)

    # Label values inside each colored segment
    for i in range(len(dtmin_vals)):
        mid1 = capex_m[i] / 2
        ax_stack.text(x_stack[i], mid1, f'{capex_m[i]:.3f} M\u20ac',
                ha='center', va='center', fontsize=9, color='white',
                fontweight='bold')
        mid2 = capex_m[i] + opex_m[i] / 2
        ax_stack.text(x_stack[i], mid2, f'{opex_m[i]:.3f} M\u20ac',
                ha='center', va='center', fontsize=9, color='white',
                fontweight='bold')

    # Total annualized cost on top of each bar
    for i, total in enumerate(total_m):
        ax_stack.text(x_stack[i], total + np.max(total_m)*0.01, f'{total:.3f} M\u20ac/yr',
                ha='center', va='bottom', fontsize=10,
                fontweight='bold', color='black')

    # Horizontal dotted line from Y-axis at the optimum (minimum) cost
    min_cost_m = min(total_m)
    opt_dtmin = dtmin_vals[np.argmin(total_m)]
    ax_stack.axhline(y=min_cost_m, color='black', linestyle=':', linewidth=1.8,
               label=f'Optimum: {min_cost_m:.3f} M€/yr @ {opt_dtmin} °C')

    ax_stack.set_xticks(x_stack)
    ax_stack.set_xticklabels([str(v) for v in dtmin_vals])
    ax_stack.set_xlabel("ΔTmin (°C)", fontsize=12)
    ax_stack.set_ylabel("Annualized Total Cost (Million €/year)", fontsize=12)
    ax_stack.legend(loc='lower left', fontsize=11, framealpha=0.9)
    ax_stack.grid(True, axis='y', alpha=0.3, linestyle='--')
    fig_stack.tight_layout()
    plt.savefig("total_cost_stacked.png", dpi=150, bbox_inches='tight')
 
 
    # Convert NPV to millions for cleaner y-axis
    npv_million = [abs(v) / 1e6 for v in npv_vals]

    # Best NPV = least-negative = smallest absolute value
    best_npv     = max(npv_vals)
    best_npv_dtmin = dtmin_vals[npv_vals.index(best_npv)]
    best_npv_million = abs(best_npv) / 1e6

    # ── Read manual calculation data from Excel ──────────────────────
    manual_file = r"C:\Users\AZ400_ATLANTIC\Desktop\Book1.xlsx"
    try:
        df_manual = pd.read_excel(manual_file)
        df_manual["DelTmin(0C)"] = pd.to_numeric(df_manual["DelTmin(0C)"], errors='coerce')
        df_manual["NPV(-Million Euro)"] = pd.to_numeric(df_manual["NPV(-Million Euro)"], errors='coerce')
        df_manual["HOC(MJ/kg of CO2)"] = pd.to_numeric(df_manual["HOC(MJ/kg of CO2)"], errors='coerce')
        df_manual.dropna(subset=["DelTmin(0C)", "NPV(-Million Euro)", "HOC(MJ/kg of CO2)"], inplace=True)
        manual_dtmin = df_manual["DelTmin(0C)"].tolist()
        manual_npv   = df_manual["NPV(-Million Euro)"].tolist()   
        manual_hoc   = df_manual["HOC(MJ/kg of CO2)"].tolist()
        has_manual = True
        print(f"Manual data loaded from {manual_file} ({len(manual_dtmin)} points).")
    except Exception as e:
        print(f"WARNING: Could not load manual data from {manual_file}: {e}")
        print("         Plotting Python results only.")
        has_manual = False
        manual_dtmin, manual_npv, manual_hoc = [], [], []

    # ── Colors ────────────────────────────────────────────────────────
    color_npv_py   = '#ff7f0e'   # Blue – Python NPV
    color_npv_man  = '#c31e23'   # Teal – Manual NPV
    color_hoc_py   = '#3594cc'   # Amber – Python HOC bars
    color_hoc_man  = '#8cc5e3'   # Light amber – Manual HOC bars

    fig_npv, ax_npv = plt.subplots(figsize=(12, 7))

    # ── Right axis: HOC (MJ/kg CO₂) — drawn first so NPV lines appear on top ──
    ax_hoc = ax_npv.twinx()
    # Send the HOC (bar) axis behind the NPV line axis
    ax_npv.set_zorder(ax_hoc.get_zorder() + 1)
    ax_npv.patch.set_visible(False)   # keep ax_npv background transparent

    bar_width = 0.35
    x_py = np.array(dtmin_vals, dtype=float)

    # Python HOC bars (shifted left)
    bars_py = ax_hoc.bar(
        x_py - bar_width / 2, hoc_vals, width=bar_width, color=color_hoc_py,
        alpha=0.6, zorder=2, label="HOC-automatic(python)"
    )
    # Manual HOC bars (shifted right)
    bars_man = None
    if has_manual:
        x_man = np.array(manual_dtmin, dtype=float)
        bars_man = ax_hoc.bar(
            x_man + bar_width / 2, manual_hoc, width=bar_width, color=color_hoc_man,
            alpha=0.6, zorder=2, label="HOC-manual"
        )
    ax_hoc.set_ylabel("HOC (MJ/kg of CO₂)", fontsize=14, color='black')
    ax_hoc.tick_params(axis='y', labelcolor='black', labelsize=14)

    all_hoc = hoc_vals + (manual_hoc if has_manual else [])
    ax_hoc.set_ylim(bottom=3, top=max(all_hoc) * 1.2)

    # ── Left axis: NPV (Million EUR) — plotted after HOC so they render on top ──
    lns1 = ax_npv.plot(
        dtmin_vals, npv_million, marker='D', linestyle='-', color=color_npv_py,
        linewidth=2.5, alpha=1.0, zorder=5, label="NPV-automatic(python)"
    )
    lns_manual_npv = []
    if has_manual:
        lns2 = ax_npv.plot(
            manual_dtmin, manual_npv, marker='o', linestyle='-', color=color_npv_man,
            linewidth=2.5, alpha=1.0, zorder=5, label="NPV-manual"
        )
        lns_manual_npv = lns2

    # Bold point for least-negative Python NPV
    sc_npv = ax_npv.scatter(
        best_npv_dtmin, best_npv_million,
        color='#d62728', s=200, zorder=6, edgecolors='black', linewidths=1.5,
        label=f"Optimal ΔTmin — Minimum NPV ({int(best_npv_dtmin)}°C)"
    )
    ax_npv.set_xlabel("ΔTmin (°C)", fontsize=14, color='black')
    ax_npv.set_ylabel("NPV (−Million €)", fontsize=14, color='black')
    ax_npv.tick_params(axis='y', labelcolor='black', labelsize=14)
    ax_npv.tick_params(axis='x', labelcolor='black', labelsize=14)
    ax_npv.set_xticks(dtmin_vals)
    ax_npv.grid(True, linestyle='--', alpha=0.5)

    # ── Combined legend ──────────────────────────────────────────────
    lns_all = lns1 + lns_manual_npv + [bars_py] + ([bars_man] if bars_man else []) + [sc_npv]
    labs_all = [l.get_label() for l in lns_all]
    ax_npv.legend(lns_all, labs_all, loc='upper center',
                  bbox_to_anchor=(0.5, -0.13), ncol=2, fontsize=9)

    # Title removed per user request
    plt.tight_layout()
    plt.savefig("npv_vs_dtmin.png", dpi=150, bbox_inches='tight')
 
    print("Plots saved: cost_vs_dtmin.png, total_cost_stacked.png, " 
          "npv_vs_dtmin.png  (HOC shown on secondary axis of NPV plot)") 
    plt.show()
 
# Main functions 
 
def main(): 
    if not os.path.exists(HYSYS_CASE_PATH): 
        print(f"ERROR: HYSYS case not found at: {HYSYS_CASE_PATH}") 
        return 
 
    # Initialise COM apartment for this thread (fixes RPC_E_SYS_CALL_FAILED)
    pythoncom.CoInitialize()
    print("Connecting to HYSYS Application...") 
    try:
        hysys = win32.Dispatch("HYSYS.Application")
    except Exception as e:
        print(f"FATAL ERROR: Could not connect to HYSYS. Make sure HYSYS is installed and open.")
        print(f"  Detail: {e}")
        return
    hysys.Visible = True 
    case = None 
 
    print(f"Opening/Finding HYSYS case: {os.path.basename(HYSYS_CASE_PATH)}...") 
    try: 
        try: 
            case = hysys.ActiveSimulationCase 
            if case and case.Pathname.lower() == HYSYS_CASE_PATH.lower(): 
                print("Using already active HYSYS case.") 
            else: 
                case = None 
        except Exception: 
            case = None 
 
        if not case: 
            for i in range(hysys.SimulationCases.Count): 
                try: 
                    c = hysys.SimulationCases.Item(i) 
                    if c.Pathname.lower() == HYSYS_CASE_PATH.lower(): 
                        case = c 
                        print(f"Found case in open cases (index {i}).") 
                        break 
                except Exception: 
                    pass 
 
        if not case: 
            print("Opening case file via SimulationCases.Open(path)...") 
            case = hysys.SimulationCases.Open(HYSYS_CASE_PATH) 
            print("Case opened successfully.") 
 
    except Exception as e: 
        print(f"FATAL ERROR during case access: {e}") 
        return 
 
    flowsheet = case.Flowsheet 
    print("Accessing collections (MaterialStreams, EnergyStreams, Operations)...") 
    ops = flowsheet.Operations 
    streams = flowsheet.MaterialStreams 
    energy_streams = flowsheet.EnergyStreams 
 
    try: 
        hx = ops.Item(HX_NAME) 
        e101 = ops.Item(AMINE_COOLER_NAME) 
        p100 = ops.Item(PUMP_100_NAME) 
        p101 = ops.Item(PUMP_101_NAME) 
        adj2 = ops.Item(DELTMIN_ADJUST_NAME) 
 
        co2_stream = streams.Item(CO2_PRODUCT_STREAM_NAME) 
        reb_energy = energy_streams.Item(REBOILER_ENERGY_STREAM_NAME) 
        cond_energy = energy_streams.Item(CONDENSER_ENERGY_STREAM_NAME) 
        print("All required objects found.") 
    except Exception as e: 
        print(f"FATAL ERROR: Could not find one or more objects: {e}") 
        return 
 
    results = [] 
    print(f"\nStarting optimization loop for Delta Tmin = {DELTMIN_VALUES}.") 
 
    for dtmin in DELTMIN_VALUES: 
        print(f"\n---> Evaluating Delta Tmin = {dtmin}.") 
 
        set_ok = False 
        candidates = ["TargetValue", "Target", "Goal", "SetPoint", "SpecValue"] 
        for candidate in candidates: 
            if hasattr(adj2, candidate): 
                try: 
                    obj = getattr(adj2, candidate) 
                    if hasattr(obj, "Value"): 
                        obj.Value = float(dtmin) 
                        set_ok = True 
                    elif isinstance(obj, (float, int)): 
                        setattr(adj2, candidate, float(dtmin)) 
                        set_ok = True 
                    if set_ok: 
                        break 
                except Exception: 
                    pass 
 
        if not set_ok and hasattr(adj2, "TargetValueValue"): 
            try: 
                adj2.TargetValueValue = float(dtmin) 
                set_ok = True 
            except Exception: 
                pass 
 
        if not set_ok: 
            raise RuntimeError( 
                f"Could not set Delta Tmin on {DELTMIN_ADJUST_NAME}. " 
                "Open the ADJ object in HYSYS and check which property holds the target." 
            ) 
 
        converged = wait_for_convergence(case, max_wait_s=MAX_WAIT_S, poll_s=POLL_S) 
        time.sleep(EXTRA_SETTLE_S) 
 
        Q_kjph = get_hx_duty_kjph(hx) 
        Q_Kw = Q_kjph / 3600.0 
        LMTD_K = get_hx_lmtd(hx) 
 
        Q_reb_kjph = get_energy_kjph(reb_energy) 
        co2_kgph = get_component_massflow_kgph(co2_stream, "CO2") 
 
        hoc_mj_per_kg = (Q_reb_kjph / co2_kgph / 1000.0) if co2_kgph > 0 else float("nan") 
 
        area_m2 = calc_area_m2(Q_Kw, U_KW_PER_M2K, LMTD_K) 
        e100_capex_2026 = calc_capex_2026_eur(area_m2) 
        reboiler_annual_opex = calc_annual_opex_eur_per_year(Q_reb_kjph) 
 
        p100_capacity_m3ph = get_pump_capacity_m3ph(p100) 
        p100_power_kw = get_pump_power_kw(p100) 
        p100_purchase_2021 = calc_pump_purchase_cost_2021(p100_capacity_m3ph) 
        p100_capex_2026 = calc_total_installed_cost_2026_from_purchase_2021( 
            p100_purchase_2021, P100_INSTALL_FACTOR 
        ) 
        p100_annual_opex = calc_pump_annual_opex_eur_per_year(p100_power_kw) 
 
        p101_capacity_m3ph = get_pump_capacity_m3ph(p101) 
        p101_power_kw = get_pump_power_kw(p101) 
        p101_purchase_2021 = calc_pump_purchase_cost_2021(p101_capacity_m3ph) 
        p101_capex_2026 = calc_total_installed_cost_2026_from_purchase_2021( 
            p101_purchase_2021, P101_INSTALL_FACTOR 
        ) 
        p101_annual_opex = calc_pump_annual_opex_eur_per_year(p101_power_kw) 
 
        Q_cond_kjph = get_energy_kjph(cond_energy) 
        Q_cond_kw = Q_cond_kjph / 3600.0 
        cond_area_m2 = calc_area_m2(Q_cond_kw, CONDENSER_U_KW_M2K, 
CONDENSER_LMTD_K) 
        cond_purchase_2021 = calc_area_based_purchase_cost_2021(cond_area_m2) 
        condenser_capex_2026 = calc_total_installed_cost_2026_from_purchase_2021( 
            cond_purchase_2021, CONDENSER_INSTALL_FACTOR 
        ) 
        condenser_annual_opex = 0.0 
 
        reb_area_m2 = calc_area_m2(Q_reb_kjph / 3600.0, REBOILER_U_KW_M2K, 
REBOILER_LMTD_K) 
        reb_purchase_2021 = calc_area_based_purchase_cost_2021(reb_area_m2) 
        reboiler_capex_2026 = calc_total_installed_cost_2026_from_purchase_2021( 
            reb_purchase_2021, REBOILER_INSTALL_FACTOR 
        ) 
 
        Q_cooler_kjph = get_hx_duty_kjph(e101) 
        Q_cooler_kw = Q_cooler_kjph / 3600.0 
        cooler_area_m2 = calc_area_m2(Q_cooler_kw, COOLER_U_KW_M2K, 
COOLER_LMTD_K) 
        cooler_purchase_2021 = calc_area_based_purchase_cost_2021(cooler_area_m2) 
        amine_cooler_capex_2026 = calc_total_installed_cost_2026_from_purchase_2021( 
            cooler_purchase_2021, COOLER_INSTALL_FACTOR 
        ) 
        amine_cooler_annual_opex = 0.0 
 
        absorber_capex_2026 = calc_direct_column_capex_2026( 
            ABSORBER_PURCHASE_COST_2021_EUR, ABSORBER_INSTALL_FACTOR 
        ) 
        desorber_capex_2026 = calc_direct_column_capex_2026( 
            DESORBER_PURCHASE_COST_2021_EUR, DESORBER_INSTALL_FACTOR 
        ) 
 
# Total CAPEX calculation 
        total_capex_2026 = ( 
            absorber_capex_2026 
            + desorber_capex_2026 
            + e100_capex_2026 
            + p100_capex_2026 
            + p101_capex_2026 
            + condenser_capex_2026 
            + reboiler_capex_2026 
            + amine_cooler_capex_2026 
        ) 
 
        maintenance_cost = MAINTENANCE_FRACTION * total_capex_2026 
 
        total_annual_opex = ( 
            p100_annual_opex 
            + p101_annual_opex 
            + reboiler_annual_opex 
            + condenser_annual_opex 
            + amine_cooler_annual_opex 
            + maintenance_cost 
            + MEA_SOLVENT_COST_EUR_PER_YEAR 
        ) 
 
        annualized_capex = total_capex_2026 / ANNUALIZATION_FACTOR_AF 
        total_annual_cost = annualized_capex + total_annual_opex 
 
        co2_tpy = calc_annual_co2_tpy(co2_kgph) 
        co2_cost_eur_per_t = total_annual_cost / co2_tpy if co2_tpy > 0 else float("inf") 
 
        row = { 
            "DelTmin_C": dtmin, 
            "Converged": bool(converged), 
 
            "HX_Q_kW": Q_Kw, 
            "HX_LMTD_K": LMTD_K, 
            "HX_Area_m2": area_m2, 
            "ReboilerDuty_kJ_per_h": Q_reb_kjph, 
            "CO2_Specific_kg_per_h": co2_kgph, 
            "CO2_t_per_y": co2_tpy, 
            "HOC_MJ_per_kg_CO2": hoc_mj_per_kg, 
 
            # E-100 
            "E100_CAPEX_2026_EUR": e100_capex_2026, 
 
            # Pump 100 
            "P100_Capacity_m3_per_h": p100_capacity_m3ph, 
            "P100_Power_kW": p100_power_kw, 
            "P100_CAPEX_2026_EUR": p100_capex_2026, 
            "P100_AnnualOPEX_EUR_per_y": p100_annual_opex, 
 
            # Pump 101 
            "P101_Capacity_m3_per_h": p101_capacity_m3ph, 
            "P101_Power_kW": p101_power_kw, 
            "P101_CAPEX_2026_EUR": p101_capex_2026, 
            "P101_AnnualOPEX_EUR_per_y": p101_annual_opex, 
 
            # Condenser 
            "CondenserDuty_kJ_per_h": Q_cond_kjph, 
            "Condenser_Area_m2": cond_area_m2, 
            "Condenser_CAPEX_2026_EUR": condenser_capex_2026, 
            "Condenser_AnnualOPEX_EUR_per_y": condenser_annual_opex, 
 
            # Reboiler 
            "Reboiler_Area_m2": reb_area_m2, 
            "Reboiler_CAPEX_2026_EUR": reboiler_capex_2026, 
            "Reboiler_AnnualOPEX_EUR_per_y": reboiler_annual_opex, 
 
            # Amine cooler 
            "AmineCoolerDuty_kJ_per_h": Q_cooler_kjph, 
            "AmineCooler_Area_m2": cooler_area_m2, 
            "AmineCooler_CAPEX_2026_EUR": amine_cooler_capex_2026, 
            "AmineCooler_AnnualOPEX_EUR_per_y": amine_cooler_annual_opex, 
 
            # Columns 
            "Absorber_CAPEX_2026_EUR": absorber_capex_2026, 
            "Desorber_CAPEX_2026_EUR": desorber_capex_2026, 
 
            # Total  
            "MaintenanceCost_EUR_per_y": maintenance_cost, 
            "MEA_SolventCost_EUR_per_y": MEA_SOLVENT_COST_EUR_PER_YEAR, 
            "CAPEX_2026_EUR": total_capex_2026, 
            "AnnualOPEX_EUR_per_y": total_annual_opex, 
            "AnnualizedCAPEX_EUR_per_y": annualized_capex, 
            "TotalAnnualCost_EUR_per_y": total_annual_cost, 
            "CO2_Captured_Cost_EUR_per_t": co2_cost_eur_per_t, 
 
            # NPV (25 yr, 7.5%): NPV = -CAPEX - OPEX * PV_annuity_factor 
            "NPV_PV_Annuity_Factor": NPV_PV_ANNUITY_FACTOR, 
            "NPV_EUR": -total_capex_2026 - total_annual_opex * NPV_PV_ANNUITY_FACTOR, 
        } 
        results.append(row) 
 
        npv = row["NPV_EUR"] 
        print( 
            f"\n{'─'*70}\n" 
            f"  Delta Tmin = {dtmin} degC  |  Converged = {converged}\n" 
            f"{'─'*70}" 
        ) 
        # ── CAPEX breakdown ──────────────────────────────────────────────── 
        print(f"  CAPEX BREAKDOWN (2026 EUR, installed):") 
        print(f"    E-100 Lean/Rich HX          : {e100_capex_2026:>15,.0f}  EUR") 
        print(f"      Area                       : {area_m2:>12,.1f}  m2") 
        print(f"    Condenser                   : {condenser_capex_2026:>15,.0f}  EUR") 
        print(f"      Area                       : {cond_area_m2:>12,.1f}  m2") 
        print(f"    Reboiler                    : {reboiler_capex_2026:>15,.0f}  EUR") 
        print(f"      Area                       : {reb_area_m2:>12,.1f}  m2") 
        print(f"    Amine Cooler                : {amine_cooler_capex_2026:>15,.0f}  EUR") 
        print(f"      Area                       : {cooler_area_m2:>12,.1f}  m2") 
        print(f"    Pump 100                    : {p100_capex_2026:>15,.0f}  EUR") 
        print(f"      Capacity                   : {p100_capacity_m3ph:>12,.1f}  m3/h") 
        print(f"    Pump 101                    : {p101_capex_2026:>15,.0f}  EUR") 
        print(f"      Capacity                   : {p101_capacity_m3ph:>12,.1f}  m3/h") 
        print(f"    Absorber (fixed)            : {absorber_capex_2026:>15,.0f}  EUR") 
        print(f"    Desorber (fixed)            : {desorber_capex_2026:>15,.0f}  EUR") 
        print(f"    {'─'*52}") 
        print(f"    TOTAL CAPEX                 : {total_capex_2026:>15,.0f}  EUR") 
        # ── OPEX breakdown ───────────────────────────────────────────────── 
        print(f"\n  OPEX BREAKDOWN (EUR/yr):") 
        print(f"    Reboiler steam              : {reboiler_annual_opex:>15,.0f}  EUR/yr") 
        print(f"      Duty                       : {Q_reb_kjph/3600.0:>12,.1f}  kW") 
        print(f"    Pump 100 electricity        : {p100_annual_opex:>15,.0f}  EUR/yr") 
        print(f"      Power                      : {p100_power_kw:>12,.2f}  kW") 
        print(f"    Pump 101 electricity        : {p101_annual_opex:>15,.0f}  EUR/yr") 
        print(f"      Power                      : {p101_power_kw:>12,.2f}  kW") 
        print(f"    Maintenance (5% x CAPEX)    : {maintenance_cost:>15,.0f}  EUR/yr") 
        print(f"    MEA solvent makeup          : {MEA_SOLVENT_COST_EUR_PER_YEAR:>15,.0f}  EUR/yr") 
        print(f"    {'─'*52}") 
        print(f"    TOTAL ANNUAL OPEX           : {total_annual_opex:>15,.0f}  EUR/yr") 
        # ── Cost summary ─────────────────────────────────────────────────── 
        print(f"\n  COST SUMMARY:") 
        print(f"    Annualised CAPEX (÷{ANNUALIZATION_FACTOR_AF})    : {annualized_capex:>15,.0f}  EUR/yr") 
        print(f"    Total Annual OPEX           : {total_annual_opex:>15,.0f}  EUR/yr") 
        print(f"    Total Annual Cost           : {annualized_capex+total_annual_opex:>15,.0f}  EUR/yr") 
        print(f"    CO2 captured                : {co2_tpy:>15,.1f}  t/yr") 
        print(f"    HOC                         : {hoc_mj_per_kg:>15.4f}  MJ/kg CO2") 
        print(f"  ► CAPTURE COST               : {co2_cost_eur_per_t:>15.3f}  EUR/tCO2") 
        print(f"    NPV                         : {npv:>15,.0f}  EUR") 
        print(f"{'─'*70}") 
 
    fieldnames = list(results[0].keys()) if results else [] 
    with open(OUT_CSV, "w", newline="", encoding="utf-8") as f: 
        w = csv.DictWriter(f, fieldnames=fieldnames) 
        w.writeheader() 
        w.writerows(results) 
 
    plot_results(results) 
 
    conv = [r for r in results if r["Converged"]] 
    pool = conv if conv else results 
 
    # Sort by ΔTmin (ascending) so the cost trend is visible left to right 
    pool_sorted = sorted(pool, key=lambda r: r["DelTmin_C"]) 
 
    best = min(pool_sorted, key=lambda r: r["CO2_Captured_Cost_EUR_per_t"]) 
    best_dtmin = best["DelTmin_C"] 
 
    # --- U-Shaped Trend Table --- 
    print("\n" + "═" * 90) 
    print(" U-SHAPED COST CURVE — CO2 CAPTURE COST vs. ΔTmin  (sorted low → high ΔTmin)") 
    print("═" * 90) 
    print(f"  {'ΔTmin(°C)':<11} {'Cost(EUR/tCO2)':<18} {'Trend':<8}" 
          f" {'Ann.CAPEX(EUR/y)':<20} {'Ann.OPEX(EUR/y)':<20} {'HX Area(m2)':<12}") 
    print("-" * 90) 
 
    costs_sorted = [r["CO2_Captured_Cost_EUR_per_t"] for r in pool_sorted] 
    for idx, r in enumerate(pool_sorted): 
        dt   = r["DelTmin_C"] 
        cost = r["CO2_Captured_Cost_EUR_per_t"] 
 
        # Trend arrow: compare with next and previous neighbour 
        if idx == 0: 
            trend = "  →" 
        else: 
            prev_cost = costs_sorted[idx - 1] 
            if cost < prev_cost: 
                trend = "  ↓" if dt != best_dtmin else "  ↓" 
            elif cost > prev_cost: 
                trend = "  ↑" 
            else: 
                trend = "  →" 
 
        # Mark the global minimum 
        is_best = (dt == best_dtmin) 
        marker  = " ◄ OPTIMAL" if is_best else "" 
        print( 
            f"  {dt:<11.1f} {cost:<18.3f} {trend:<8}" 
            f" {r['AnnualizedCAPEX_EUR_per_y']:<20,.0f}" 
            f" {r['AnnualOPEX_EUR_per_y']:<20,.0f}" 
            f" {r['HX_Area_m2']:<12.1f}" 
            f"{marker}" 
        ) 
 
    print("-" * 90) 
    print() 
    print(f"  ↓ = Cost decreasing as ΔTmin increases") 
    print(f"  ↑ = Cost increasing as ΔTmin increases (CAPEX penalty dominates at low ΔTmin,") 
    print(f"      or OPEX penalty dominates at high ΔTmin)") 
    print(f"  ◄ OPTIMAL = Global minimum capture cost (sweet spot)") 
 
    # --- Detect sub-regions: where does descent start, where does ascent start --- 
    ascending_after_min  = [r["DelTmin_C"] for r in pool_sorted 
                            if r["DelTmin_C"] > best_dtmin 
                            and r["CO2_Captured_Cost_EUR_per_t"] > 
best["CO2_Captured_Cost_EUR_per_t"]] 
    descending_before_min = [r["DelTmin_C"] for r in pool_sorted 
                             if r["DelTmin_C"] < best_dtmin 
                             and r["CO2_Captured_Cost_EUR_per_t"] > 
best["CO2_Captured_Cost_EUR_per_t"]] 
 
    print() 
    print("  COST TREND ZONES (based on neighbours):") 
    if descending_before_min: 
        print(f"   • High ΔTmin zone (OPEX-dominated, cost rises):  " 
              f">= {ascending_after_min[0]:.1f} °C" if ascending_after_min else "   • (no high-cost zone detected above optimum)") 
        print(f"   • Descending zone (cost falls toward optimum):   " 
              f"{min(descending_before_min):.1f} – {best_dtmin:.1f} °C") 
    print(f"   • OPTIMAL ΔTmin: {best_dtmin:.1f} °C  →  " 
          f"{best['CO2_Captured_Cost_EUR_per_t']:.3f} EUR/tCO2") 
    if ascending_after_min: 
        print(f"   • Ascending zone (CAPEX grows too large):        " 
              f"{ascending_after_min[0]:.1f} – {ascending_after_min[-1]:.1f} °C") 
 
    print() 
    print("  OPTIMAL POINT DETAIL:") 
    print(f"    Delta Tmin        : {best['DelTmin_C']} °C") 
    print(f"    Capture Cost      : {best['CO2_Captured_Cost_EUR_per_t']:.3f} EUR/tCO2") 
    print(f"    Total CAPEX 2026  : {best['CAPEX_2026_EUR']:,.0f} EUR") 
    print(f"    Annualized CAPEX  : {best['AnnualizedCAPEX_EUR_per_y']:,.0f} EUR/y") 
    print(f"    Total Annual OPEX : {best['AnnualOPEX_EUR_per_y']:,.0f} EUR/y") 
    print(f"    HOC               : {best['HOC_MJ_per_kg_CO2']:.4f} MJ/kg CO2") 
    print(f"    E-100 HX Area     : {best['HX_Area_m2']:.1f} m²") 
    print(f"    CO2 Captured      : {best['CO2_t_per_y']:,.0f} t/yr") 
    print("═" * 90) 
 
    # NPV summary – least-negative NPV is the optimum 
    best_npv_row = max(pool_sorted, key=lambda r: r["NPV_EUR"]) 
    print("\n===== NPV ANALYSIS (25 yr, 7.5% discount rate) =====") 
    print(f"PV Annuity Factor : {NPV_PV_ANNUITY_FACTOR:.4f}") 
    print(f"Formula           : NPV = -CAPEX - OPEX × {NPV_PV_ANNUITY_FACTOR:.4f}") 
    print() 
    print(f"{'ΔTmin (°C)':<12} {'CAPEX (EUR)':>18} {'Annual OPEX (EUR/y)':>22} {'NPV (EUR)':>20}") 
    print("-" * 76) 
    for r in pool_sorted: 
        marker = "  <-- BEST NPV" if r["DelTmin_C"] == best_npv_row["DelTmin_C"] else "" 
        print( 
            f"{r['DelTmin_C']:<12} " 
            f"{r['CAPEX_2026_EUR']:>18,.0f} " 
            f"{r['AnnualOPEX_EUR_per_y']:>22,.0f} " 
            f"{r['NPV_EUR']:>20,.0f}" 
            f"{marker}" 
        ) 
    print("-" * 76) 
    print(f"\n>>> Optimal ΔTmin (least-negative NPV): {best_npv_row['DelTmin_C']} °C") 
    print(f"    CAPEX      : {best_npv_row['CAPEX_2026_EUR']:,.0f} EUR") 
    print(f"    Annual OPEX: {best_npv_row['AnnualOPEX_EUR_per_y']:,.0f} EUR/y") 
    print(f"    NPV        : {best_npv_row['NPV_EUR']:,.0f} EUR") 
 
if __name__ == "__main__": 
    main()
    print(f"\nSaved results to: {os.path.abspath(OUT_CSV)}")
    print("\nNOTE:")
    print("CEPCI_2021 was not present in the provided PDFs.")
    print("Please verify/update CEPCI_2021 before using the final economic numbers in reporting.") 