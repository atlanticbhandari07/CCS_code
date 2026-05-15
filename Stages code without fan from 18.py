import os 
import time 
import math 
import csv 
from datetime import datetime 
import win32com.client as win32 
import matplotlib.pyplot as plt 
# ========================================================= 
# USER SETTINGS 
# ========================================================= 
HYSYS_CASE_PATH = r"D:\BUUU\working file without fan.hsc"
ABSORBER_NAME = "Absorber"                    
# Unit op name in HYSYS 
CO2_PRODUCT_STREAM_NAME = "CO2 product"       # CO2 product stream 
REBOILER_ENERGY_STREAM_NAME = "Reboiler duty" # Reboiler duty energy stream 
CONDENSER_ENERGY_STREAM_NAME = "Condenser duty"  # Condenser duty stream 
CLEAN_GAS_STREAM_NAME = "Clean gas"           
# Clean gas outlet stream 
# Capture efficiency: inlet CO2 mole fraction (fixed) 
INLET_CO2_MOLE_FRACTION = 0.1546 
# Additional equipment names in HYSYS 
HX_NAME = "E-100"                             
# Main Lean/Rich HEX 
AMINE_COOLER_NAME = "E-101"                   
PUMP_100_NAME = "P-100"                       
# Amine Cooler 
# Pump 100 
PUMP_101_NAME = "P-101"                       
# Pump 101 
# Column stages / packing height range to optimize 
STAGE_VALUES = list(range(18, 26, 1))  
# CO2 component efficiency to set on all stages after each stage increment
# (default in HYSYS is 1.0; must be reset to 0.15 after adding new stages)
CO2_STAGE_EFFICIENCY = 0.15

# Base case for Absorber CAPEX scaling 
BASE_PACKING_HEIGHT = 18 
BASE_COST_SS_2022_EUR = 1127200.0             
EXPONENT = 0.9 
# CAPEX constants 
SS_TO_CS_FACTOR = 1.75 
INSTALL_FACTOR = 5.3 
CEPCI_2021 = 820.97 
CEPCI_2022 = 816.0 
CEPCI_2026 = 830.0 
# E-100 HEX cost parameters 
U_KW_PER_M2K = 1.5 
# Base purchased cost (SS) in 2022 EUR 
C1_SS_2021_EUR = 1662000.0     # SS316 purchased cost (2021) at reference area A1_M2 
A1_M2 = 2747.0                 
EXPONENT_E = 1.0               
# Reference area (m²) for base cost C1_SS_2021_EUR 
# Scaling exponent (1.0 = linear with area) 
E100_INSTALL_FACTOR = 5.3      # E-100 installation factor 
# HEX area-based reference cost (CS, 2021) for Condenser, Reboiler, Amine Cooler 
HEX_AREA_REF_COST_2021_EUR = C1_SS_2021_EUR / SS_TO_CS_FACTOR 
# Condenser parameters 
CONDENSER_U_KW_M2K = 1.0 
CONDENSER_LMTD_K = 71.14 
CONDENSER_INSTALL_FACTOR = 9.480 
# Reboiler parameters 
REBOILER_U_KW_M2K = 1.2 
REBOILER_LMTD_K = 27.18 
REBOILER_INSTALL_FACTOR = 5.30 
# Amine Cooler parameters 
COOLER_U_KW_M2K = 0.8 
COOLER_LMTD_K = 27.0 
COOLER_INSTALL_FACTOR = 5.30 
# Pump cost parameters 
PUMP_REF_CAPACITY_M3PH = 940.0 
PUMP_REF_COST_2021_EUR = 42307.7 
PUMP_COST_EXPONENT = 0.6 
P100_INSTALL_FACTOR = 7.580 
P101_INSTALL_FACTOR = 8.90 
# Desorber column cost 
DESORBER_PURCHASE_COST_2021_EUR = 4385258.0 
DESORBER_INSTALL_FACTOR = 3.330 
# OPEX constants 
OPERATING_HOURS_PER_YEAR = 8000.0 
STEAM_PRICE_EUR_PER_KWH = 0.015 
ELECTRICITY_PRICE_EUR_PER_KWH = 0.05 
ANNUALIZATION_FACTOR_AF = 11.15 
MAINTENANCE_FRACTION = 0.05 
MEA_SOLVENT_COST_EUR_PER_YEAR = 708571.0 
# Names of Streams for E-100 Heat Exchanger (Lean/Rich) 
LEAN_HOT_IN_NAME = "Lean MEA 2" 
LEAN_HOT_OUT_NAME = "Lean MEA 3" 
RICH_COLD_IN_NAME = "Rich MEA 2" 
RICH_COLD_OUT_NAME = "Rich MEA 3" 
# Solver / timing 
MAX_WAIT_S = 300  # 5 minutes — increased for HYSYS convergence
POLL_S = 1.0 
EXTRA_SETTLE_S = 1.0 
# Output 
OUT_CSV =f"absorber_stage_optimization_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv" 
# ========================================================= 
# HELPER FUNCTIONS 
# ========================================================= 
def wait_for_convergence(case, max_wait_s=300, poll_s=2.0): 
    """ 
    Wait until HYSYS solver finishes. 
    """ 
    start_time = time.time() 
    print("  Waiting for HYSYS solver...", end="", flush=True) 
 
    solver = None 
    try: 
        solver = case.Solver 
    except Exception: 
        pass 
 
    time.sleep(1.0) 
 
    while time.time() - start_time < max_wait_s: 
        is_solving = True 
        try: 
            if solver is not None: 
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
 
def force_run(case, absorber=None):
    """
    Force HYSYS to run the absorber column after a structural change.
    Equivalent to pressing the 'Run' button on the column sub-flowsheet.
    """
    # Step 1 — Run the column sub-flowsheet (= pressing the Run button in HYSYS)
    if absorber is not None:
        try:
            absorber.ColumnFlowsheet.Run()
            print("  [force_run] absorber.ColumnFlowsheet.Run() called.")
        except Exception as e:
            print(f"  [force_run] ColumnFlowsheet.Run() failed: {e}")
            try:
                absorber.Run()
                print("  [force_run] absorber.Run() called (fallback).")
            except Exception as e2:
                print(f"  [force_run] absorber.Run() also failed: {e2}")

    # Step 2 — Toggle main solver to propagate changes through the flowsheet
    try:
        case.Solver.CanSolve = False
        time.sleep(0.5)
        case.Solver.CanSolve = True
        print("  [force_run] Main solver toggled ON.")
    except Exception as e:
        print(f"  [force_run] Main solver toggle failed: {e}")
 
def wait_for_column_convergence(absorber, case, max_wait_s=180, poll_s=2.0):
    """
    Wait for the absorber COLUMN sub-flowsheet to converge.
    Polls ColumnFlowsheet.Converged — this is the green/red status bar in HYSYS.
    """
    start = time.time()
    print("  Waiting for absorber column to converge...", end="", flush=True)

    while time.time() - start < max_wait_s:
        # Primary: ColumnFlowsheet.Converged flag (green bar = True in HYSYS)
        try:
            if absorber.ColumnFlowsheet.Converged:
                print(" Column converged.")
                return True
        except Exception:
            pass

        # Secondary: main solver idle — assume column is done
        try:
            if not case.Solver.IsSolving:
                time.sleep(poll_s)
                try:
                    if absorber.ColumnFlowsheet.Converged:
                        print(" Column converged (main solver idle + Converged).")
                        return True
                except Exception:
                    pass
                print(" Main solver idle — assuming column done.")
                return True
        except Exception:
            pass

        print(".", end="", flush=True)
        time.sleep(poll_s)

    print(" Timeout waiting for column convergence.")
    return False

 
def get_component_massflow_kgph(stream, comp_name="CO2"): 
    """ 
    Read component mass flow (kg/h) from a HYSYS material stream. 
    """ 
    try: 
        fp = stream.FluidPackage 
        comps = fp.Components 
        comp_idx = -1 
 
        for i in range(comps.Count): 
            try: 
                cname = comps.Item(i).Name 
            except Exception: 
                cname = comps.Item(i).name 
            if str(cname).strip().upper() == comp_name.upper(): 
                comp_idx = i 
                break 
 
        if comp_idx == -1:
            raise RuntimeError(
                f"Component '{comp_name}' not found in fluid package for stream '{stream.Name}'."
            ) 
 
        if hasattr(stream, "ComponentMassFlowValue"): 
            cmf = stream.ComponentMassFlowValue 
            if len(cmf) > comp_idx: 
                return float(cmf[comp_idx]) * 3600.0 
 
        total_kgph = 0.0 
        mf = stream.MassFlow 
        if hasattr(mf, "GetValue"): 
            total_kgph = float(mf.GetValue("kg/h")) 
        elif hasattr(mf, "Value"): 
            total_kgph = float(mf.Value) * 3600.0 
        else: 
            total_kgph = float(mf) * 3600.0 
 
        co2_frac = 0.0 
        if hasattr(stream, "ComponentMassFractionValue"): 
            cmfv = stream.ComponentMassFractionValue 
            if len(cmfv) > comp_idx: 
                co2_frac = float(cmfv[comp_idx]) 
 
        if total_kgph > 0 and co2_frac > 0: 
            return total_kgph * co2_frac 
 
    except Exception as e: 
        print(f"  Warning: component CO2 mass flow read failed: {e}") 
 
    raise RuntimeError(
        f"Could not read component mass flow for '{comp_name}' from stream '{stream.Name}'."
    ) 
 
def get_energy_kjph(energy_stream): 
    """ 
    Read energy flow in kJ/h from energy stream. 
    """ 
    candidates = ["HeatFlow", "Duty", "EnergyFlow", "Q"] 
 
    for c in candidates: 
        if hasattr(energy_stream, c): 
            v = getattr(energy_stream, c) 
            try: 
                if hasattr(v, "GetValue"): 
                    return float(v.GetValue("kJ/h")) 
                if hasattr(v, "Value"): 
                    return float(v.Value) 
                return float(v) 
            except Exception: 
                pass 
 
    raise RuntimeError("Could not read energy stream heat flow in kJ/h.") 
 
def get_co2_mole_fraction(stream, comp_name="CO2"): 
    """ 
    Read CO2 mole fraction from a HYSYS material stream. 
    Tries ComponentMoleFractionValue array first, then MolarComposition. 
    """ 
    try: 
        fp = stream.FluidPackage 
        comps = fp.Components 
        comp_idx = -1 
        for i in range(comps.Count): 
            try: 
                cname = comps.Item(i).Name 
            except Exception: 
                cname = comps.Item(i).name 
            if str(cname).strip().upper() == comp_name.upper(): 
                comp_idx = i 
                break 
 
        if comp_idx == -1: 
            raise RuntimeError(f"Component '{comp_name}' not found in fluid package for stream '{stream.Name}'.") 
 
        # Try ComponentMoleFractionValue (tuple of all component mole fractions) 
        for attr_name in ["ComponentMoleFractionValue", "ComponentMolarFractionValue"]: 
            if hasattr(stream, attr_name): 
                cmfv = getattr(stream, attr_name) 
                if len(cmfv) > comp_idx: 
                    val = float(cmfv[comp_idx]) 
                    if val >= 0: 
                        return val 
 
        # Fallback: MolarComposition or MoleFraction property 
        for attr_name in ["MolarComposition", "MoleFraction", "ComponentMoleFraction"]: 
            if hasattr(stream, attr_name): 
                prop = getattr(stream, attr_name) 
                if hasattr(prop, "GetValue"): 
                    try: 
                        val = float(prop.GetValue("")) 
                        return val 
                    except Exception: 
                        pass 
 
    except Exception as e: 
        print(f"  Warning: CO2 mole fraction read failed: {e}") 
 
    raise RuntimeError(f"Could not read CO2 mole fraction from stream '{stream.Name}'.") 
 
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
 
 
 
 
 
def _set_absorber_col1_stages(absorber, n_stages):
    """
    Set the number of stages on the absorber column's COL1
    packing/tray section.
    """
    target = int(n_stages)
    errors = []

    STAGE_PROPS = [
        "NumberOfStages", "NumStages", "StageCount",
        "NumberOfTrays", "TrayCount",
        "NumberOfPacks", "PackCount",
    ]

    # Path 1: ColumnFlowsheet.Operations -> find COL1/MAIN TS/MAIN TOWER
    try:
        cf = absorber.ColumnFlowsheet
        ops = cf.Operations
        for i in range(ops.Count):
            try:
                sub_op = ops.Item(i)
                sub_name = ""
                try:
                    sub_name = sub_op.name
                except Exception:
                    try:
                        sub_name = sub_op.Name
                    except Exception:
                        sub_name = f"<op_{i}>"
                
                if sub_name.upper() in ["COL1", "MAIN TS", "MAIN TOWER"]:
                    for prop_name in STAGE_PROPS:
                        if hasattr(sub_op, prop_name):
                            prop = getattr(sub_op, prop_name)
                            try:
                                if hasattr(prop, "SetValue"):
                                    prop.SetValue(target, "")
                                    return
                                if hasattr(prop, "Value"):
                                    prop.Value = target
                                    return
                                setattr(sub_op, prop_name, target)
                                return
                            except Exception as e:
                                errors.append(f"{sub_name}.{prop_name}: {e}")
            except Exception as e:
                errors.append(f"Operations.Item({i}): {e}")
    except Exception as e:
        errors.append(f"ColumnFlowsheet.Operations: {e}")

    raise RuntimeError(
        f"Could not set Absorber COL1 number of stages to {target}. "
        f"Details: {errors}"
    )
 
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
 
def get_pump_capacity_m3ph(pump): 
    """ 
    Read pump volumetric capacity in m3/h. 
    """ 
    candidates = [ 
        "Capacity", "capacity", 
        "PumpCapacity", "RatedCapacity", "VolumetricFlow", "LiquidFlow", 
    ] 
    units = ["m3/h", "m3/hr", "m^3/h", "m^3/hr"] 
    errors = [] 
    for name in candidates: 
        if hasattr(pump, name): 
            prop = getattr(pump, name) 
            try: 
                return _try_get_value(prop, units) 
            except Exception as e: 
                errors.append(f"{name}: {e}") 
 
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
    """ 
    units_kw   = ["kW"] 
    units_kjph = ["kJ/h"] 
    errors = [] 
 
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
 
# ========================================================= 
# COST CALCULATION FUNCTIONS 
# ========================================================= 
def calc_area_m2(Q_Kw, U_Kw_m2k, LMTD_K): 
    """A = Q / (U * LMTD).""" 
    if U_Kw_m2k <= 0 or LMTD_K <= 0: 
        return float("nan") 
    return Q_Kw / (U_Kw_m2k * LMTD_K) 
 
def escalate_cost_2021_to_2026(cost_2021): 
    return cost_2021 * (CEPCI_2026 / CEPCI_2021) 
 
def calc_e100_capex_2026_eur(area_m2): 
    """ 
    E-100 (lean/rich HX) installed CAPEX escalated to 2026. 
    1. Scale purchased SS316 cost (2021) with area 
    2. Convert SS316 -> CS (material factor) 
    3. Apply installation factor 
    4. Escalate 2021 -> 2026 via CEPCI 
    """ 
    if area_m2 <= 0: 
        return float("nan") 
    C2_SS_2021 = C1_SS_2021_EUR * ((area_m2 / A1_M2) ** EXPONENT_E) 
    C2_CS_2021 = C2_SS_2021 / SS_TO_CS_FACTOR 
    capex_2021 = C2_CS_2021 * E100_INSTALL_FACTOR 
    capex_2026 = capex_2021 * (CEPCI_2026 / CEPCI_2021) 
    return capex_2026 
 
def calc_area_based_purchase_cost_2021(area_m2, 
    ref_cost_2021=None): 
    if ref_cost_2021 is None: 
        ref_cost_2021 = HEX_AREA_REF_COST_2021_EUR 
    if area_m2 <= 0: 
        return float("nan") 
    return ref_cost_2021 * ((area_m2 / A1_M2) ** EXPONENT_E) 
 
def calc_total_installed_cost_2026_from_purchase_2021(purchase_cost_2021, 
    install_factor): 
    if purchase_cost_2021 <= 0: 
        return float("nan") 
    return escalate_cost_2021_to_2026(purchase_cost_2021 * install_factor) 
 
def calc_absorber_capex_2026_eur(stages): 
    """ 
    Absorber CAPEX based on stage/packing-height scaling: 
        C_new_SS_2022 = C_base_SS_2022 * (stages / base_stages)^0.9 
        C_new_CS_2022 = C_new_SS_2022 / 1.75 
        CAPEX_2022 = C_new_CS_2022 * 5.3 
        CAPEX_2026 = CAPEX_2022 * (830 / 816) 
    """ 
    if stages <= 0: 
        return float("nan") 
    c_ss_2022 = BASE_COST_SS_2022_EUR * ((float(stages) / BASE_PACKING_HEIGHT) ** EXPONENT) 
    c_cs_2022 = c_ss_2022 / SS_TO_CS_FACTOR 
    capex_2022 = c_cs_2022 * INSTALL_FACTOR 
    capex_2026 = capex_2022 * (CEPCI_2026 / CEPCI_2022) 
    return capex_2026 
 
def calc_pump_purchase_cost_2021(capacity_m3ph): 
    if capacity_m3ph <= 0: 
        return float("nan") 
    return PUMP_REF_COST_2021_EUR * ( 
        (capacity_m3ph / PUMP_REF_CAPACITY_M3PH) ** PUMP_COST_EXPONENT 
    ) 
 
def calc_pump_annual_opex_eur_per_year(power_kw):
    if power_kw < 0:
        return float("nan")
    return power_kw * OPERATING_HOURS_PER_YEAR * ELECTRICITY_PRICE_EUR_PER_KWH 
 
def calc_direct_column_capex_2026(purchase_cost_2021, install_factor): 
    return calc_total_installed_cost_2026_from_purchase_2021( 
        purchase_cost_2021, install_factor 
    ) 
 
def calc_annual_reboiler_opex_eur_per_year(q_reb_kjph): 
    """ 
    Annual OPEX = (kJ/h / 3600 -> kW) * operating hours * steam price 
    """ 
    q_reb_kw = q_reb_kjph / 3600.0 
    annual_energy_kwh = q_reb_kw * OPERATING_HOURS_PER_YEAR 
    return annual_energy_kwh * STEAM_PRICE_EUR_PER_KWH 
 
 
 
 
 
def calc_annual_co2_tpy(co2_kgph): 
    """Convert kg/h to t/y using 8000 h/y.""" 
    return (co2_kgph * OPERATING_HOURS_PER_YEAR) / 1000.0 
 
# ========================================================= 
# STAGE-SETTING FUNCTIONS 
# ========================================================= 
def set_absorber_stages(absorber, case, n_stages):
    """
    Set the number of stages on the absorber column's COL1 sub-flowsheet.
    Targets: absorber.ColumnFlowsheet.Operations.Item(i).NumberOfStages
    which is the 'Num of Stages' field in HYSYS Design > Connections tab.
    """
    target = int(n_stages)
    errors = []

    # Pause main solver before structural change
    old_main = True
    try:
        old_main = case.Solver.CanSolve
        case.Solver.CanSolve = False
        print("    [set_stages] Main solver paused.")
    except Exception as e:
        print(f"    [set_stages] Could not pause main solver (continuing): {e}")

    STAGE_PROPS = [
        "NumberOfStages", "NumStages", "StageCount",
        "NumberOfTrays", "TrayCount",
        "NumberOfPacks", "PackCount",
    ]

    def _try_set_prop(obj, prop, label):
        """Try all three write strategies for a COM property."""
        if not hasattr(obj, prop):
            return False
        try:
            attr = getattr(obj, prop)
            if hasattr(attr, "Value"):
                attr.Value = target
                return True
        except Exception as e:
            errors.append(f"{label}.{prop}.Value: {e}")
        try:
            attr = getattr(obj, prop)
            if hasattr(attr, "SetValue"):
                attr.SetValue(target, "")
                return True
        except Exception as e:
            errors.append(f"{label}.{prop}.SetValue: {e}")
        try:
            setattr(obj, prop, target)
            return True
        except Exception as e:
            errors.append(f"{label}.{prop} direct: {e}")
        return False

    set_ok = False
    set_msg = ""

    try:
        cf = absorber.ColumnFlowsheet
        ops = cf.Operations
        n_ops = int(ops.Count)
        print(f"    [set_stages] Scanning {n_ops} ColumnFlowsheet sub-operation(s)...")

        for i in range(n_ops):
            try:
                sub_op = ops.Item(i)
                sub_name = ""
                try:
                    sub_name = sub_op.name
                except Exception:
                    try:
                        sub_name = sub_op.Name
                    except Exception:
                        sub_name = f"<op_{i}>"

                print(f"    [set_stages]   Sub-op [{i}]: '{sub_name}'")
                
                # Check if name is one of our targets
                if sub_name.upper() in ["COL1", "MAIN TS", "MAIN TOWER"]:
                    # Try direct stage props on this sub-op
                    for prop in STAGE_PROPS:
                        if _try_set_prop(sub_op, prop, f"ops[{i}]({sub_name})"):
                            set_msg = f"Path C: ops[{i}]({sub_name}).{prop} = {target}"
                            print(f"    [set_stages] {set_msg}")
                            set_ok = True
                            break

                    if set_ok:
                        break

                    # Try sub-containers within this sub-op
                    for inner in ("TraySection", "PackingSection", "Section",
                                   "Design", "Parameters", "Specs"):
                        if not hasattr(sub_op, inner):
                            continue
                        try:
                            inner_obj = getattr(sub_op, inner)
                            for prop in STAGE_PROPS:
                                if _try_set_prop(inner_obj, prop,
                                                 f"ops[{i}]({sub_name}).{inner}"):
                                    set_msg = (f"Path C2: ops[{i}]({sub_name})"
                                               f".{inner}.{prop} = {target}")
                                    print(f"    [set_stages] {set_msg}")
                                    set_ok = True
                                    break
                            if set_ok:
                                break
                        except Exception as e:
                            errors.append(f"ops[{i}].{inner}: {e}")
                    if set_ok:
                        break

            except Exception as e:
                errors.append(f"ops.Item({i}): {e}")

    except Exception as e:
        errors.append(f"ColumnFlowsheet.Operations: {e}")

    # Always restore main solver
    try:
        case.Solver.CanSolve = old_main
        print("    [set_stages] Main solver restored.")
        
        # Reset and Run column to force structural change to propagate
        try:
            cf = absorber.ColumnFlowsheet
            cf.Reset()
            print("    [set_stages] ColumnFlowsheet.Reset() called.")
        except Exception:
            pass
    except Exception:
        pass

    if not set_ok:
        raise RuntimeError(
            f"set_absorber_stages: could not set COL1 stages to {target}.\n"
            + "\n".join(f"  \u2022 {err}" for err in errors)
        )

    return set_msg
 

 
def get_absorber_stages(absorber):
    """
    Read back the current number of stages from the absorber's COL1
    sub-flowsheet operation — the same path used by set_absorber_stages.
    """
    STAGE_PROPS = [
        "NumberOfStages", "NumStages", "StageCount",
        "NumberOfTrays", "TrayCount",
    ]

    # Primary: read from ColumnFlowsheet.Operations sub-ops
    try:
        cf = absorber.ColumnFlowsheet
        ops = cf.Operations
        for i in range(int(ops.Count)):
            try:
                sub_op = ops.Item(i)
                sub_name = ""
                try:
                    sub_name = sub_op.name
                except Exception:
                    try:
                        sub_name = sub_op.Name
                    except Exception:
                        sub_name = ""
                
                if sub_name.upper() in ["COL1", "MAIN TS", "MAIN TOWER"] or not sub_name:
                    for prop in STAGE_PROPS:
                        if hasattr(sub_op, prop):
                            try:
                                val = getattr(sub_op, prop)
                                if hasattr(val, "Value"):
                                    return int(float(val.Value))
                                return int(float(val))
                            except Exception:
                                pass
            except Exception:
                pass
    except Exception:
        pass

    # Fallback: ColumnStages.Count
    try:
        cf = absorber.ColumnFlowsheet
        if hasattr(cf, "ColumnStages"):
            return int(cf.ColumnStages.Count)
    except Exception:
        pass

    return float("nan")
 


# =========================================================
# CO2 STAGE EFFICIENCY SETTING
# =========================================================
def set_co2_stage_efficiency(absorber, co2_eff=0.15, comp_name="CO2"):
    """
    Set the CO2 component stage efficiency = co2_eff for ALL stages in COL1.

    This corresponds to Parameters > Efficiencies > Component table in HYSYS,
    where each stage row shows: Nitrogen=1.0, CO2=0.150, H2O=1.0, MEAmine=1.0 ...

    When you add a new stage (e.g. going from 18→19), HYSYS inserts it with
    CO2 efficiency = 1.0 by default. This function resets it to 0.15.

    Returns: (stages_set, total_stages) — how many stages were updated.
    """
    errors = []

    try:
        cf = absorber.ColumnFlowsheet
        ops = cf.Operations

        # ── Step 1: find the column section sub-op ──────────────────────
        # Don't filter by name — the sub-op may not be called "COL1".
        # Instead pick the first sub-op that has a stage-count property,
        # which is the same strategy used by set_absorber_stages().
        col1 = None
        col1_name = ""
        STAGE_COUNT_PROPS = ("NumberOfStages", "NumStages", "StageCount",
                             "NumberOfTrays", "TrayCount")
        for i in range(int(ops.Count)):
            try:
                sub_op = ops.Item(i)
                try:
                    sub_name = sub_op.name
                except Exception:
                    try:
                        sub_name = sub_op.Name
                    except Exception:
                        sub_name = f"<op_{i}>"
                
                if sub_name.upper() in ["COL1", "MAIN TS", "MAIN TOWER"] or any(hasattr(sub_op, p) for p in STAGE_COUNT_PROPS):
                    col1 = sub_op
                    col1_name = sub_name
                    print(f"  [eff] Using column section sub-op [{i}]: '{sub_name}'")
                    break
            except Exception as e:
                errors.append(f"ops.Item({i}): {e}")

        if col1 is None:
            # Nothing found — dump all sub-op names to help diagnose
            diag = []
            try:
                for i in range(int(ops.Count)):
                    try:
                        n = ops.Item(i).Name
                    except Exception:
                        n = f"<op_{i}>"
                    diag.append(f"[{i}]={n}")
            except Exception:
                pass
            raise RuntimeError(
                f"No column section sub-op found in ColumnFlowsheet.Operations "
                f"({int(ops.Count)} sub-op(s): {diag}). Errors: {errors}"
            )


        # ── Step 2: find CO2 component index ────────────────────────────
        co2_idx = -1
        try:
            fp = cf.FluidPackage
        except Exception:
            try:
                fp = absorber.FluidPackage
            except Exception as e:
                raise RuntimeError(f"Cannot access FluidPackage: {e}")

        comps = fp.Components
        for i in range(int(comps.Count)):
            try:
                cname = comps.Item(i).Name
            except Exception:
                try:
                    cname = comps.Item(i).name
                except Exception:
                    cname = ""
            if str(cname).strip().upper() == comp_name.upper():
                co2_idx = i
                break

        if co2_idx == -1:
            raise RuntimeError(
                f"Component '{comp_name}' not found in fluid package "
                f"(checked {int(comps.Count)} components)."
            )

        print(f"  [eff] CO2 is component index {co2_idx} in COL1.")

        # ── Step 3: get stage count ──────────────────────────────────────
        n_stages = 0
        for prop in ("NumberOfStages", "NumStages", "StageCount",
                     "NumberOfTrays", "TrayCount"):
            if hasattr(col1, prop):
                try:
                    val = getattr(col1, prop)
                    n_stages = int(float(val.Value) if hasattr(val, "Value")
                                   else float(val))
                    break
                except Exception:
                    pass

        if n_stages == 0:
            raise RuntimeError("Could not read COL1 NumberOfStages.")

        print(f"  [eff] Setting CO2 efficiency = {co2_eff} for {n_stages} stages...")

        # ── Step 4: set efficiency for every stage ───────────────────────
        set_count = 0

        # --- Path 1: ComponentEfficiencies 2-D array on col1 ---
        # HYSYS 14 typically exposes this as a (n_stages × n_comps) SafeArray.
        for arr_attr in ("ComponentEfficiencies", "ComponentEfficiency",
                         "StageEfficiencies", "EfficiencyValues", "TrayEfficiencies",
                         "MolarEfficiency", "ComponentMolarEfficiency", "SectionEfficiencies"):
            if hasattr(col1, arr_attr):
                try:
                    eff_array = getattr(col1, arr_attr)
                    for s in range(n_stages):
                        try:
                            eff_array[s][co2_idx] = co2_eff
                            set_count += 1
                        except Exception:
                            pass
                    if set_count == n_stages:
                        print(f"  [eff] Path 1 ({arr_attr}): set {set_count}/{n_stages} stages.")
                        return set_count, n_stages
                    # partial success — still useful, keep count
                except Exception as e:
                    errors.append(f"Path1 {arr_attr}: {e}")

        # --- Path 2: per-stage objects ---
        for stages_attr in ("Stages", "TrayStages", "ColumnStages", "Trays", "Tray"):
            if hasattr(col1, stages_attr):
                try:
                    stages_coll = getattr(col1, stages_attr)
                    prev_count = set_count
                    n_coll = int(stages_coll.Count)
                    for s in range(n_coll):
                        try:
                            stage_obj = stages_coll.Item(s)
                            # Try multiple candidate names for efficiency
                            for eff_attr in ("ComponentEfficiency", "Efficiency", 
                                             "TrayEfficiency", "StageEfficiency",
                                             "ComponentEfficiencies", "MolarEfficiency",
                                             "ComponentMolarEfficiency"):
                                if hasattr(stage_obj, eff_attr):
                                    try:
                                        ea = getattr(stage_obj, eff_attr)
                                        # If it's a property/collection, try to set by index
                                        try:
                                            ea[co2_idx] = co2_eff
                                            set_count += 1
                                            break
                                        except Exception:
                                            if hasattr(ea, "Item"):
                                                ea.Item(co2_idx).Value = co2_eff
                                                set_count += 1
                                                break
                                            if hasattr(ea, "SetValue"):
                                                ea.SetValue(co2_idx, co2_eff)
                                                set_count += 1
                                                break
                                    except Exception:
                                        pass
                        except Exception as e:
                            errors.append(f"Path2 stage[{s}]: {e}")
                    if set_count > prev_count:
                        print(f"  [eff] Path 2 ({stages_attr}): set {set_count-prev_count}/{n_stages} stages.")
                        if set_count >= n_stages:
                            return set_count, n_stages
                except Exception as e:
                    errors.append(f"Path2 {stages_attr}: {e}")

        # --- Path 3: SetComponentEfficiency(stage, comp, value) method ---
        if hasattr(col1, "SetComponentEfficiency"):
            try:
                prev_count = set_count
                for s in range(n_stages):
                    try:
                        col1.SetComponentEfficiency(s, co2_idx, co2_eff)
                        set_count += 1
                    except Exception as e:
                        errors.append(f"Path3 SetComponentEfficiency(s={s}): {e}")
                if set_count > prev_count:
                    print(f"  [eff] Path 3 (SetComponentEfficiency): "
                          f"set {set_count - prev_count}/{n_stages} stages.")
            except Exception as e:
                errors.append(f"Path3: {e}")

        if set_count == 0:
            raise RuntimeError(
                f"Could not set CO2 efficiency on any stage. "
                f"Details: {errors}"
            )

        if set_count < n_stages:
            print(f"  [eff] WARNING: only set {set_count}/{n_stages} stages.")

        return set_count, n_stages

    except RuntimeError:
        raise
    except Exception as e:
        raise RuntimeError(f"set_co2_stage_efficiency: {e}")

# ========================================================= 
# PLOTTING 
# ========================================================= 
def plot_results(results): 
    """ 
    Generate plots for: 
    1. CO2 capture cost vs stages 
    2. Annualized CAPEX vs stages 
    3. Annual OPEX vs stages 
    """ 
    if not results: 
        return 
 
    x = [r["ColumnStages"] for r in results] 
    cost = [r["CO2_Captured_Cost_EUR_per_t"] for r in results] 
    capex = [r["AnnualizedCAPEX_EUR_per_y"] for r in results] 
    opex = [r["AnnualOPEX_EUR_per_y"] for r in results] 
 
    # Plot 1 
    plt.figure(figsize=(11, 6)) 
    plt.plot(x, cost, marker="o", linewidth=2, label="$CO_2$ Capture Cost") 
    best_idx = cost.index(min(cost)) 
    plt.scatter([x[best_idx]], [cost[best_idx]], s=120, zorder=5, label="Minimum cost Point") 
    plt.xlabel("Absorber stages") 
    plt.ylabel(r"$CO_2$ Captured Cost (€/t $CO_2$)") 
    plt.xticks(x) 
    plt.grid(True, linestyle="--", alpha=0.7) 
    plt.legend() 
    plt.savefig("absorber_cost_vs_stages.png", dpi=150, bbox_inches="tight") 
 
    # Plot 2 
    plt.figure(figsize=(10, 6)) 
    plt.plot(x, capex, marker="s", linewidth=2, label="Annualized CAPEX") 
    plt.xlabel("Absorber Column Stages") 
    plt.ylabel("Annualized CAPEX (EUR / year)") 
    plt.title("Annualized CAPEX vs Absorber Column Stages") 
    plt.xticks(x) 
    plt.grid(True, linestyle="--", alpha=0.7) 
    plt.legend() 
    plt.savefig("absorber_capex_vs_stages.png", dpi=150, bbox_inches="tight") 
 
    # Plot 3 
    plt.figure(figsize=(10, 6)) 
    plt.plot(x, opex, marker="^", linewidth=2, label="Annual OPEX") 
    plt.xlabel("Absorber Column Stages") 
    plt.ylabel("Annual OPEX (EUR / year)") 
    plt.title("Annual OPEX vs Absorber Column Stages") 
    plt.xticks(x) 
    plt.grid(True, linestyle="--", alpha=0.7) 
    plt.legend() 
    plt.savefig("absorber_opex_vs_stages.png", dpi=150, bbox_inches="tight") 
 
    print("Plots saved:") 
    print("  absorber_cost_vs_stages.png") 
    print("  absorber_capex_vs_stages.png") 
    print("  absorber_opex_vs_stages.png") 
    plt.show()
 
# ========================================================= 
# MAIN 
# ========================================================= 
def main(): 
    if not os.path.exists(HYSYS_CASE_PATH): 
        print(f"ERROR: HYSYS case not found at: {HYSYS_CASE_PATH}") 
        return 
 
    print("Connecting to HYSYS Application...") 
    hysys = win32.Dispatch("HYSYS.Application") 
    hysys.Visible = True 
 
    case = None 
    print(f"Opening/Finding HYSYS case: {os.path.basename(HYSYS_CASE_PATH)} ...") 
 
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
        absorber = ops.Item(ABSORBER_NAME) 
        hx = ops.Item(HX_NAME) 
        e101 = ops.Item(AMINE_COOLER_NAME) 
        p100 = ops.Item(PUMP_100_NAME) 
        p101 = ops.Item(PUMP_101_NAME) 
 
        co2_stream = streams.Item(CO2_PRODUCT_STREAM_NAME) 
        clean_gas_stream = streams.Item(CLEAN_GAS_STREAM_NAME) 
        reb_energy = energy_streams.Item(REBOILER_ENERGY_STREAM_NAME) 
        cond_energy = energy_streams.Item(CONDENSER_ENERGY_STREAM_NAME) 
 
        print("All required objects found.") 
    except Exception as e: 
        print(f"FATAL ERROR: Could not find one or more HYSYS objects: {e}") 
        return 
 
    results = [] 
    print(f"\nStarting absorber optimization loop for stages = {STAGE_VALUES}") 
 
    for n_stages in STAGE_VALUES: 
        print(f"\n---> Evaluating absorber column stages = {n_stages}") 

        # ── Base case: 18 stages already converged, skip stage-change & Run ──
        is_base_case = (n_stages == STAGE_VALUES[0])
        if is_base_case:
            print(f"  Base case ({n_stages} stages) — already converged, skipping stage change.")
            converged = True
            col_converged = True
        else:
            # ── Step 1: Change the stage count ───────────────────────────────
            try:
                prop_used = set_absorber_stages(absorber, case, n_stages)
                print(f"  Set absorber stages using property: {prop_used}")
            except Exception as e:
                print(f"  ERROR while setting stages: {e}")
                continue

            # ── Step 2: Set CO2 efficiency = 0.15 on ALL stages (incl. new) ──
            # Must be done BEFORE Run so the new stage doesn't solve with eff=1.
            try:
                n_set, n_total = set_co2_stage_efficiency(
                    absorber, co2_eff=CO2_STAGE_EFFICIENCY, comp_name="CO2"
                )
                print(f"  CO2 efficiency set to {CO2_STAGE_EFFICIENCY} "
                      f"on {n_set}/{n_total} stages.")
            except Exception as e:
                print(f"  WARNING: Could not set CO2 stage efficiency: {e}")
                print(f"  Continuing — efficiency may be wrong for new stage(s).")

            # ── Step 3: Press Run (ColumnFlowsheet.Run) + wait ───────────────
            force_run(case, absorber)
            col_converged = wait_for_column_convergence(
                absorber, case, max_wait_s=MAX_WAIT_S, poll_s=POLL_S
            )
            converged = wait_for_convergence(case, max_wait_s=MAX_WAIT_S, poll_s=POLL_S)
            time.sleep(EXTRA_SETTLE_S)
            if not col_converged:
                print(f"  WARNING: absorber column did not converge for {n_stages} stages — results may be unreliable.")


 
        # ── Read process data ────────────────────────────────────────── 
        try: 
            q_reb_kjph = get_energy_kjph(reb_energy) 
            co2_kgph = get_component_massflow_kgph(co2_stream, "CO2") 
            actual_stages = get_absorber_stages(absorber) 
 
            # E-100 HEX 
            Q_e100_kjph = get_hx_duty_kjph(hx) 
            Q_e100_kw = Q_e100_kjph / 3600.0 
            LMTD_e100 = get_hx_lmtd(hx) 
 
            # Pump 100 
            p100_capacity_m3ph = get_pump_capacity_m3ph(p100) 
            p100_power_kw = get_pump_power_kw(p100) 
 
            # Pump 101 
            p101_capacity_m3ph = get_pump_capacity_m3ph(p101) 
            p101_power_kw = get_pump_power_kw(p101) 
 
            # Condenser 
            Q_cond_kjph = get_energy_kjph(cond_energy) 
            Q_cond_kw = Q_cond_kjph / 3600.0 
 
            # Amine Cooler (E-101) 
            Q_cooler_kjph = get_hx_duty_kjph(e101) 
            Q_cooler_kw = Q_cooler_kjph / 3600.0 
 
            # Clean gas CO2 mole fraction (for capture efficiency) 
 
            clean_gas_co2_molefrac = get_co2_mole_fraction(clean_gas_stream, "CO2") 
 
        except Exception as e: 
            print(f"  ERROR while reading HYSYS data: {e}") 
            continue 
 
        # ── HOC ──────────────────────────────────────────────────────── 
        hoc_mj_per_kg = (q_reb_kjph / co2_kgph / 1000.0) if co2_kgph > 0 else float("nan") 
 
        # ── Capture Efficiency ─────────────────────────────────────────
        if INLET_CO2_MOLE_FRACTION > 0:
            capture_efficiency = (
                (INLET_CO2_MOLE_FRACTION - clean_gas_co2_molefrac)
                / INLET_CO2_MOLE_FRACTION
            )
        else:
            capture_efficiency = float("nan")
        capture_efficiency_pct = capture_efficiency * 100.0 
 
        # ── CAPEX calculations ───────────────────────────────────────── 
        # Absorber (stage-scaled) 
        absorber_capex_2026 = calc_absorber_capex_2026_eur(n_stages) 
 
        # E-100 HEX (area-scaled) 
        e100_area_m2 = calc_area_m2(Q_e100_kw, U_KW_PER_M2K, LMTD_e100) 
        e100_capex_2026 = calc_e100_capex_2026_eur(e100_area_m2) 
 
        # Pump 100 
        p100_purchase_2021 = calc_pump_purchase_cost_2021(p100_capacity_m3ph) 
        p100_capex_2026 = calc_total_installed_cost_2026_from_purchase_2021( 
            p100_purchase_2021, P100_INSTALL_FACTOR 
        ) 
        p100_annual_opex = calc_pump_annual_opex_eur_per_year(p100_power_kw) 
 
        # Pump 101 
        p101_purchase_2021 = calc_pump_purchase_cost_2021(p101_capacity_m3ph) 
        p101_capex_2026 = calc_total_installed_cost_2026_from_purchase_2021( 
            p101_purchase_2021, P101_INSTALL_FACTOR 
        ) 
        p101_annual_opex = calc_pump_annual_opex_eur_per_year(p101_power_kw) 
 
        # Condenser 
        cond_area_m2 = calc_area_m2(Q_cond_kw, CONDENSER_U_KW_M2K, CONDENSER_LMTD_K) 
        cond_purchase_2021 = calc_area_based_purchase_cost_2021(cond_area_m2) 
        condenser_capex_2026 = calc_total_installed_cost_2026_from_purchase_2021( 
            cond_purchase_2021, CONDENSER_INSTALL_FACTOR 
        ) 
        condenser_annual_opex = 0.0 
 
        # Reboiler 
        reb_area_m2 = calc_area_m2(q_reb_kjph / 3600.0, REBOILER_U_KW_M2K, REBOILER_LMTD_K) 
        reb_purchase_2021 = calc_area_based_purchase_cost_2021(reb_area_m2) 
        reboiler_capex_2026 = calc_total_installed_cost_2026_from_purchase_2021( 
            reb_purchase_2021, REBOILER_INSTALL_FACTOR 
        ) 
        reboiler_annual_opex = calc_annual_reboiler_opex_eur_per_year(q_reb_kjph) 
 
        # Amine Cooler (E-101) 
        cooler_area_m2 = calc_area_m2(Q_cooler_kw, COOLER_U_KW_M2K, COOLER_LMTD_K) 
        cooler_purchase_2021 = calc_area_based_purchase_cost_2021(cooler_area_m2) 
        amine_cooler_capex_2026 = calc_total_installed_cost_2026_from_purchase_2021( 
            cooler_purchase_2021, COOLER_INSTALL_FACTOR 
        ) 
        amine_cooler_annual_opex = 0.0 
 
        # Desorber (fixed cost) 
        desorber_capex_2026 = calc_direct_column_capex_2026( 
            DESORBER_PURCHASE_COST_2021_EUR, DESORBER_INSTALL_FACTOR 
        ) 
 
        # ── TOTALS ───────────────────────────────────────────────────── 
 
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
        co2_capture_cost = total_annual_cost / co2_tpy if co2_tpy > 0 else float("inf") 
 
        # ── Results row ──────────────────────────────────────────────── 
        row = { 
            "ColumnStages": n_stages, 
            "ActualStages_Readback": actual_stages, 
            "Converged": bool(converged), 
            "ReboilerDuty_kJ_per_h": q_reb_kjph, 
            "CO2_Component_kg_per_h": co2_kgph, 
            "CO2_t_per_y": co2_tpy, 
            "HOC_MJ_per_kg_CO2": hoc_mj_per_kg, 
            "CleanGas_CO2_MoleFrac": clean_gas_co2_molefrac, 
            "Capture_Efficiency_pct": capture_efficiency_pct, 
 
            # E-100 HEX 
            "E100_Q_kW": Q_e100_kw, 
            "E100_LMTD_K": LMTD_e100, 
            "E100_Area_m2": e100_area_m2, 
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
 
            # Amine Cooler 
            "AmineCoolerDuty_kJ_per_h": Q_cooler_kjph, 
            "AmineCooler_Area_m2": cooler_area_m2, 
            "AmineCooler_CAPEX_2026_EUR": amine_cooler_capex_2026, 
            "AmineCooler_AnnualOPEX_EUR_per_y": amine_cooler_annual_opex, 
 
            # Columns 
 
            "Absorber_CAPEX_2026_EUR": absorber_capex_2026, 
            "Desorber_CAPEX_2026_EUR": desorber_capex_2026, 
 
            # Totals 
            "MaintenanceCost_EUR_per_y": maintenance_cost, 
            "MEA_SolventCost_EUR_per_y": MEA_SOLVENT_COST_EUR_PER_YEAR, 
            "CAPEX_2026_EUR": total_capex_2026, 
            "AnnualOPEX_EUR_per_y": total_annual_opex, 
            "AnnualizedCAPEX_EUR_per_y": annualized_capex, 
            "TotalAnnualCost_EUR_per_y": total_annual_cost, 
            "CO2_Captured_Cost_EUR_per_t": co2_capture_cost, 
        } 
        results.append(row) 
 
        # ── Detailed print breakdown ─────────────────────────────────── 
        print( 
            f"\n{'─'*70}\n" 
            f"  Stages = {n_stages}  |  Converged = {converged}\n" 
            f"{'─'*70}" 
        ) 
        # CAPEX breakdown 
        print(f"  CAPEX BREAKDOWN (2026 EUR, installed):") 
        print(f"    E-100 Lean/Rich HX          : {e100_capex_2026:>15,.0f}  EUR") 
        print(f"      Area                       : {e100_area_m2:>12,.1f}  m2") 
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
        print(f"    Absorber (stage-scaled)     : {absorber_capex_2026:>15,.0f}  EUR") 
        print(f"    Desorber (fixed)            : {desorber_capex_2026:>15,.0f}  EUR") 
 
        print(f"    {'─'*52}") 
        print(f"    TOTAL CAPEX                 : {total_capex_2026:>15,.0f}  EUR") 
 
        # OPEX breakdown 
        print(f"\n  OPEX BREAKDOWN (EUR/yr):") 
        print(f"    Reboiler steam              : {reboiler_annual_opex:>15,.0f}  EUR/yr") 
        print(f"      Duty                       : {q_reb_kjph/3600.0:>12,.1f}  kW") 
        print(f"    Pump 100 electricity        : {p100_annual_opex:>15,.0f}  EUR/yr") 
        print(f"      Power                      : {p100_power_kw:>12,.2f}  kW") 
        print(f"    Pump 101 electricity        : {p101_annual_opex:>15,.0f}  EUR/yr") 
        print(f"      Power                      : {p101_power_kw:>12,.2f}  kW") 
 
        print(f"    Maintenance (5% x CAPEX)    : {maintenance_cost:>15,.0f}  EUR/yr") 
        print(f"    MEA solvent makeup          : {MEA_SOLVENT_COST_EUR_PER_YEAR:>15,.0f}  EUR/yr") 
        print(f"    {'─'*52}") 
        print(f"    TOTAL ANNUAL OPEX           : {total_annual_opex:>15,.0f}  EUR/yr") 
 
        # Cost summary 
        print(f"\n  COST SUMMARY:") 
        print(f"    Annualised CAPEX (÷{ANNUALIZATION_FACTOR_AF})    : {annualized_capex:>15,.0f}  EUR/yr") 
        print(f"    Total Annual OPEX           : {total_annual_opex:>15,.0f}  EUR/yr") 
        print(f"    Total Annual Cost           : {total_annual_cost:>15,.0f}  EUR/yr") 
        print(f"    CO2 captured                : {co2_tpy:>15,.1f}  t/yr") 
        print(f"    HOC                         : {hoc_mj_per_kg:>15.4f}  MJ/kg CO2") 
        print(f"  ► CAPTURE COST               : {co2_capture_cost:>15.3f}  EUR/tCO2") 
        print(f"    Capture Efficiency          : {capture_efficiency_pct:>15.2f}  %") 
        print(f"      Clean gas CO2 mole frac    : {clean_gas_co2_molefrac:>12.6f}") 
        print(f"      Inlet CO2 mole frac (fixed): {INLET_CO2_MOLE_FRACTION:>12.4f}") 
        print(f"{'─'*70}") 
 
    if not results: 
        print("No results generated.") 
        return 
 
    # Save CSV 
    fieldnames = list(results[0].keys()) 
    with open(OUT_CSV, "w", newline="", encoding="utf-8") as f: 
        writer = csv.DictWriter(f, fieldnames=fieldnames) 
        writer.writeheader() 
        writer.writerows(results) 
 
    # Plots 
    plot_results(results) 
 
    # Best point 
    converged_only = [r for r in results if r["Converged"]] 
    pool = converged_only if converged_only else results 
    best = min(pool, key=lambda r: r["CO2_Captured_Cost_EUR_per_t"]) 
 
    print("\n===== MINIMUM CO2 CAPTURE COST POINT =====") 
    print(f"Best Column Stages: {best['ColumnStages']}") 
    print(f"Readback Stages: {best['ActualStages_Readback']}") 
    print(f"CO2 Captured Cost: {best['CO2_Captured_Cost_EUR_per_t']:.2f} EUR/tCO2") 
    print(f"Total CAPEX 2026: {best['CAPEX_2026_EUR']:,.0f} EUR") 
    print(f"Annualized CAPEX: {best['AnnualizedCAPEX_EUR_per_y']:,.0f} EUR/y") 
    print(f"Total Annual OPEX: {best['AnnualOPEX_EUR_per_y']:,.0f} EUR/y") 
    print(f"Total Annual Cost: {best['TotalAnnualCost_EUR_per_y']:,.0f} EUR/y") 
    print(f"HOC: {best['HOC_MJ_per_kg_CO2']:.4f} MJ/kg CO2") 
    print(f"Capture Efficiency: {best['Capture_Efficiency_pct']:.2f} %") 
    print(f"Annual CO2 Captured: {best['CO2_t_per_y']:.2f} t/y") 
    print(f"\nSaved results to: {os.path.abspath(OUT_CSV)}") 
 
if __name__ == "__main__": 
    main() 