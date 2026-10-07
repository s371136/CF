# -*- coding: utf-8 -*-
"""
Simulazione oraria off-grid 

Esegui questo file con F5 in Spyder.

Versione unificata: 8 scenari economici (LOW/BASE/HIGH + 4 fonti di
letteratura + MEDIAN) e vincoli terminali SOC/LOH: SOC_end >= SOC_initial e
LOH_end >= LOH_initial.

Legge il foglio 'Input data', intervallo B6:D8765, del file Excel indicato
sotto. Le colonne devono essere: domanda [kW], capacity factor wind,
capacity factor PV. Wind e PV sono gia normalizzati nel file Excel.
"""

from pathlib import Path
import math
import tkinter as tk
from tkinter import filedialog, messagebox
import numpy as np
import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

try:
    import matplotlib.pyplot as plt
    from matplotlib import rcParams
    MATPLOTLIB_AVAILABLE = True
except ModuleNotFoundError:
    MATPLOTLIB_AVAILABLE = False


# =============================================================================
# INPUT DATA
# Tutti i dati modificabili sono raccolti qui e separati per componente.
# Potenze in kW, energie in kWh, costi in euro (valori reali all'anno 0).
# =============================================================================

# --- File di input e opzioni di esecuzione -----------------------------------
# Il file Excel viene scelto dall'utente tramite una finestra di selezione.
# Lasciare None: load_profiles() aprira automaticamente la finestra quando
# il simulatore viene eseguito direttamente o importato da un ottimizzatore.
EXCEL_FILE = None
SHEET_NAME = "Input data"
PLOT_RESULTS = True
SAVE_EXCEL_RESULTS = True   # True = apre la finestra "Salva con nome" a fine simulazione
RESULTS_EXCEL_FILE = "simulation_results.xlsx"  # nome proposto, non salvataggio automatico
# Valori proposti nella finestra di selezione grafici.
# Non limitano piu' automaticamente il plotting a Week 1 / Month 1.
PLOT_WEEK_NUMBER = 1       # valore iniziale proposto, 1-52
PLOT_MONTH_NUMBER = 1      # valore iniziale proposto, 1-12
PLOT_ANNUAL_DEFAULT = True # mostrare per default anche la vista annuale sintetica

# --- Taglie dei componenti ---------------------------------------------------
# Hornsund baseline.
# PV_SIZE = 50 kW is the estimated site upper bound and is used here as a
# provisional baseline, not as an already optimized design.
PV_SIZE = 50.0              # kW, Hornsund upper bound / provisional baseline

# Wind: numero intero di turbine (max ~5 da documento Hornsund).
# WT_RATED_POWER e' preso da un datasheet della turbina ICEWIND RW scelta
N_WT_MAX = 5                # numero massimo di turbine
WT_RATED_POWER = 0.6       # kW/turbina 
N_WT = 5                    # numero di turbine del caso baseline
WIND_SIZE = N_WT * WT_RATED_POWER   # kW

FC_SIZE = 100.0             # kW, TEMPORARY ASSUMPTION - TO OPTIMIZE
EL_SIZE = 55.0              # kW, TEMPORARY ASSUMPTION - TO OPTIMIZE
H2_STORAGE_SIZE = 3333.0    # kWh_H2 LHV, TEMPORARY ASSUMPTION - TO OPTIMIZE
BATTERY_SIZE = 500.0        # kWh, planned Hornsund battery capacity

# Hornsund has two 65.5 kW gensets, normally operated alternately.
# This baseline represents the active unit; the second unit is backup.
DIESEL_SIZE = 65.5          # kW, rated electrical power of one genset
DIESEL_UNIT_SIZE = 65.5     # kW, gruppo esistente (usato come bound dagli ottimizzatori)

# --- Scelte utente -----------------------------------------------------------
# Se lo script viene eseguito direttamente (F5 in Spyder), queste opzioni
# vengono chieste con una finestra grafica prima di selezionare il file Excel.
APPLY_HORNSUND_CONSTRAINTS = None
SYSTEM_MODE = None  # "diesel_only", "renewable_only", "mix", "battery_only_mix", "battery_only_renewable"

# Copia delle taglie di baseline, cosi' ogni modalita' parte dagli stessi valori.
_BASELINE_SIZES = {
    "PV_SIZE": PV_SIZE,
    "N_WT": N_WT,
    "FC_SIZE": FC_SIZE,
    "EL_SIZE": EL_SIZE,
    "H2_STORAGE_SIZE": H2_STORAGE_SIZE,
    "BATTERY_SIZE": BATTERY_SIZE,
    "DIESEL_SIZE": DIESEL_SIZE,
}

# --- Vincoli di progetto Hornsund --------------------------------------------
# These are site/design constraints. They do not change the hourly dispatch
# directly; they are checked here and will later become constraints of the
# sizing/optimization problem.
H2_SYSTEM_BUDGET = 250000.0          # EUR, total EL + H2 storage + FC budget
PV_MAX_SIZE = 50.0                   # kWp, massimo installabile a Hornsund

# Batteria prevista per Hornsund.
# Assunzione di scenario:
#   - con i vincoli Hornsund attivi, nei sistemi che includono una batteria
#     la capacita' e' fissata a 500 kWh;
#   - senza vincoli Hornsund, la taglia puo' essere modificata/ottimizzata
#     dall'ottimizzatore esterno.
# Il documento di sito indica ~500 kWh come sistema pianificato, non gia' installato.
BATTERY_HORNSUND_FIXED_SIZE = 500.0   # kWh
BATTERY_MAX_SIZE = BATTERY_HORNSUND_FIXED_SIZE  # mantenuto per compatibilita'

# Vincoli terminali annuali degli accumuli (latest STORAGE_END formulation).
# A design is feasible only if it does not consume net initial stored energy:
#   SOC_battery,end >= SOC_battery,initial
#   LOH_H2,end       >= LOH_H2,initial
# Small tolerances only avoid floating-point false infeasibilities.
BATTERY_SOC_END_TOLERANCE = 1e-6
H2_LOH_END_TOLERANCE = 1e-6
# Una sola consegna di gasolio l'anno (~80.000 L), di cui ~2.000 L per i
# trasporti a terra (Table 4 dal nostro file word). Ipotesi: il resto e' disponibile per l'elettricita'.
MAX_DIESEL_FUEL_L_PER_YEAR = 78000.0  # L/anno per generazione elettrica
MAX_MASS_PER_SHIPMENT_KG = 10000.0   # kg per sea shipment
MAX_MASS_PER_CRANE_ITEM_KG = 1500.0  # kg per individual item during unloading

# Component masses are not yet available. Fill these values from supplier
# datasheets when components are selected. None = constraint not yet checkable.
COMPONENT_ITEM_MASS_KG = {
    "Battery Li-ion": None,
    "PEM electrolyzer": None,
    "PEM fuel cell": None,
    "H2 storage unit": None,
    "Wind turbine unit": 80,    #kg
    "PV shipment item": None,
}

# --- Parametri finanziari generali ------------------------------------------
# Per confrontare le fonti economiche senza confondere l'effetto dei costi con
# quello dell'orizzonte finanziario, tutti gli scenari usano la stessa struttura
# NPC/LCOE armonizzata su 20 anni; il discount rate resta quello della più recente versione STORAGE_END (5%).
PROJECT_LIFETIME = 20      # years
DISCOUNT_RATE = 0.05       # fraction/year  (retained from STORAGE_END latest version)

# Reference exchange rates used for the source-cost table (05/10/2026).
# Currency conversion only: no inflation/source-year escalation is applied.
USD_TO_EUR = 0.89254       # EUR/USD
CAD_TO_EUR = 0.62621       # EUR/CAD

# --- Sistema fotovoltaico ----------------------------------------------------
PV_CAPEX = 1547.0                  # EUR/kW, overwritten by selected cost scenario
PV_REPLACEMENT_FRACTION = 1.0      # fraction of CAPEX
PV_LIFETIME = 20.0                 # years
PV_FIXED_OM = 24.0                 # EUR/(kW year)
PV_VARIABLE_OM = 0.0               # EUR/kWh generated

# --- Sistema eolico ----------------------------------------------------------
# WARNING: no source below is specific to the 0.6 kW IceWind RW600. The wind
# CAPEX scenarios are therefore literature sensitivities, not supplier quotes.
WIND_CAPEX = 1175.0                # EUR/kW, overwritten by selected cost scenario
WIND_REPLACEMENT_FRACTION = 1.0    # fraction of CAPEX
WIND_LIFETIME = 21.0               # years
WIND_FIXED_OM = 35.25              # EUR/(kW year)
WIND_VARIABLE_OM = 0.0             # EUR/kWh generated

# --- Batteria Li-ion ---------------------------------------------------------
BATTERY_SOC_INITIAL = 0.50         # fraction
BATTERY_REMAINING_LIFE_INITIAL = 1.00  # fraction
BATTERY_SOC_MIN = 0.20             # fraction
BATTERY_SOC_MAX = 1.00             # fraction
BATTERY_CHARGE_EFFICIENCY = 0.90   # includes converter efficiency
BATTERY_DISCHARGE_EFFICIENCY = 0.90  # includes converter efficiency
BATTERY_SELF_DISCHARGE = 0.0       # fraction/hour
# Economic values are overwritten by the selected literature scenario.
BATTERY_CAPEX = 550.0
BATTERY_REPLACEMENT_COST = 275.0
BATTERY_FIXED_OM = 10.0
BATTERY_VARIABLE_OM = 0.0          # EUR/kWh battery throughput
BATTERY_CTF_DOD_AVERAGE = (0.5 * 4993 + 0.7 * 2961 + 0.8 * 2487) / 3  # cycles

# --- Elettrolizzatore PEM ----------------------------------------------------
# Efficienza di SISTEMA su base LHV, dipendente dal carico parziale.
# I punti sono stati digitalizzati dalla Fig. 4b di Marocco et al. (2021),
# "Optimal design of stand-alone solutions based on RES + hydrogen storage
# feeding off-grid communities", Energy Conversion and Management 238, 114147.
# La figura riporta la curva di system efficiency (stack + BOP); il paper usa
# poi un fit polinomiale. Qui usiamo interpolazione lineare tra i punti
# digitalizzati, per evitare overshoot artificiali del fit.
EL_MIN_LOAD_FRACTION = 0.10        # fraction of rated power, paper Table 2
EL_EFFICIENCY_LOAD_FRACTION = np.array([
    0.10, 0.12, 0.15, 0.20, 0.25, 0.30, 0.40,
    0.50, 0.60, 0.70, 0.80, 0.90, 1.00
], dtype=float)
EL_EFFICIENCY_LHV = np.array([
    0.397, 0.437, 0.477, 0.512, 0.530, 0.539, 0.546,
    0.545, 0.541, 0.535, 0.529, 0.523, 0.516
], dtype=float)
EL_REMAINING_LIFE_INITIAL = 1.00   # fraction
# Economic values are overwritten by the selected literature scenario.
# In the Marocco reference case EL_CAPEX is the reference specific cost at
# 50 kW and is size-scaled by the helper function defined below.
EL_CAPEX = 4600.0                   # EUR/kW reference or constant, scenario-dependent
EL_STACK_REPLACEMENT_FRACTION = 0.267
EL_STACK_LIFETIME = 40000.0        # operating hours
EL_STACK_START_LIFETIME = 5000.0   # starts
EL_TOTAL_OM_FRACTION = 0.04        # fraction CAPEX/year (Marocco formulation)
EL_FIXED_OM_PER_KW_YEAR = 0.0      # EUR/(kW year), used by McKinley
EL_CAPEX_REFERENCE_SIZE = 50.0     # kW, Marocco reference size
EL_COST_EXPONENT = 0.65            # Marocco size-cost exponent
EL_USE_SIZE_SCALING = True

# --- Serbatoio di idrogeno ---------------------------------------------------
H2_MIN_PRESSURE = 3.0              # bar
H2_MAX_PRESSURE = 28.0             # bar
"H2_SOC_INITIAL = H2_MIN_PRESSURE/H2_MAX_PRESSURE              # fraction"
H2_SOC_INITIAL = 0.5             # fraction
H2_LHV = 119.96                    # MJ/kg
# Economic values are overwritten by the selected literature scenario.
H2_TANK_CAPEX = 470.0              # EUR/kg H2
H2_TANK_REPLACEMENT_COST = 0.0     # EUR/kg H2; Marocco does not specify a tank replacement
H2_TANK_LIFETIME = 35.0            # years
H2_TANK_FIXED_OM_FRACTION = 0.02   # fraction of tank CAPEX/year
H2_TANK_FIXED_OM_PER_KG_YEAR = 0.0 # EUR/(kg H2 year), used by McKinley

# Janke treats water treatment and compression as separate H2-system CAPEX
# tied to electrolyzer rated power. They are zero in Marocco/McKinley presets.
H2_WATER_TREATMENT_CAPEX_PER_KW_EL = 0.0  # EUR/kW_EL
H2_COMPRESSOR_CAPEX_PER_KW_EL = 0.0       # EUR/kW_EL
# Janke uses an all-encompassing H2-system OPEX per kWh charged/discharged.
H2_SYSTEM_VARIABLE_OM = 0.0        # EUR/kWh_H2 throughput

# --- Fuel cell PEM -----------------------------------------------------------
# Efficienza di SISTEMA su base LHV, dipendente dal carico parziale.
# Punti digitalizzati dalla Fig. 4c dello stesso paper.
FC_MIN_LOAD_FRACTION = 0.06        # fraction of rated power, paper Table 2
FC_EFFICIENCY_LOAD_FRACTION = np.array([
    0.06, 0.08, 0.10, 0.12, 0.15, 0.20, 0.25, 0.30,
    0.40, 0.50, 0.60, 0.70, 0.80, 0.90, 1.00
], dtype=float)
FC_EFFICIENCY_LHV = np.array([
    0.449, 0.488, 0.515, 0.534, 0.553, 0.570, 0.576, 0.578,
    0.572, 0.560, 0.543, 0.522, 0.497, 0.467, 0.427
], dtype=float)
FC_REMAINING_LIFE_INITIAL = 1.00   # fraction
# Economic values are overwritten by the selected literature scenario.
# In the Marocco reference case FC_CAPEX is the reference specific cost at
# 10 kW and is size-scaled by the helper function defined below.
FC_CAPEX = 3947.0                   # EUR/kW reference or constant, scenario-dependent
FC_STACK_REPLACEMENT_FRACTION = 0.267
FC_STACK_LIFETIME = 30000.0         # operating hours
FC_STACK_START_LIFETIME = 10000.0   # starts, Marocco Table 1/2 framework
FC_TOTAL_OM_FRACTION = 0.04         # fraction CAPEX/year (Marocco formulation)
FC_FIXED_OM_PER_KW_YEAR = 0.0       # EUR/(kW year)
FC_OM_PER_OPERATING_HOUR = 0.0      # EUR/operating hour, McKinley formulation
FC_CAPEX_REFERENCE_SIZE = 10.0      # kW, Marocco reference size
FC_COST_EXPONENT = 0.70             # Marocco size-cost exponent
FC_USE_SIZE_SCALING = True

# --- Generatore diesel -------------------------------------------------------
# DIESEL_FUEL_CURVE_A/B sono riferiti alla potenza MECCANICA
# ALL'ALBERO del motore (solo motore diesel). Il rendimento dell'alternatore
# viene applicato a parte (vedi sezione successiva). Se A e B fossero invece
# ricavati dalla potenza elettrica del gruppo completo, l'alternatore sarebbe
# contato due volte: in tal caso porre ALTERNATOR_EFFICIENCY_ENABLED = False.
DIESEL_MIN_LOAD_FRACTION = 0.30       # fraction of rated power
DIESEL_REMAINING_LIFE_INITIAL = 1.00  # fraction
DIESEL_FUEL_CURVE_A = 0.327           # L/kWh (all'albero)
DIESEL_FUEL_CURVE_B = -0.00626        # L/kWh (all'albero, per kW nominale all'albero)
DIESEL_START_DURATION = 4.0           # minutes at rated power
# Hornsund uses the existing diesel genset: initial CAPEX is therefore a sunk
# cost and is kept at zero in every cost-source scenario. Replacement and fuel
# economics are intentionally held constant across the technology-cost scenarios
# so that Arctic CAPEX/OPEX sensitivity is not confounded by a different diesel
# price assumption. A separate diesel-price sensitivity can be run later.
DIESEL_CAPEX = 0.0
DIESEL_REPLACEMENT_COST = 420.0     # EUR/kW, Marocco reference
DIESEL_LIFETIME = 20000.0           # operating hours
DIESEL_FIXED_OM_FRACTION = 0.0
DIESEL_VARIABLE_OM = 0.0            # EUR/kWh_el
DIESEL_OM_PER_OPERATING_HOUR = 0.40 # EUR/h, Marocco
DIESEL_OM_PER_L = 0.0               # EUR/L, Janke-style channel if desired
DIESEL_FUEL_COST = 2.00              # EUR/L, held constant in source scenarios
DIESEL_EMISSION_FACTOR = 2.68        # kgCO2/L, retained from STORAGE_END latest version

# =============================================================================
# UNIFIED ECONOMIC SCENARIOS: LOW / BASE / HIGH + 4 LITERATURE SOURCES + MEDIAN
# =============================================================================
# LOW / BASE / HIGH are the three Arctic/remote sensitivity cases from the
# latest STORAGE_END simulator. The four named literature presets are taken
# from the literature-cost simulator and retain its harmonised 20-year cost
# structure. The MEDIAN scenario is calculated parameter-by-parameter from the
# four named literature presets only. LOW / BASE / HIGH remain separate
# sensitivity scenarios and are not used to construct the literature median.
# All eight dictionaries expose the same keys, so the same GUI and optimizer
# can switch among them safely.
#
# IMPORTANT:
# - LOW / BASE / HIGH are sensitivity scenarios, not individual publications.
# - MAROCCO_2022 uses the published EL/FC size-scaling functions.
# - MCKINLEY/JANKE presets are harmonised sensitivity cases, not exact
#   replications of the source studies.
# - Existing Hornsund diesel initial CAPEX is treated as sunk (= 0).
# - MEDIAN is a synthetic literature-based scenario, not a publication; it is
#   recalculated from the four named literature presets every time this module
#   is imported.

COST_SCENARIO = "BASE"
COST_SCENARIO_KEY = "BASE"

COST_SCENARIO_LABELS = {
    "LOW": "LOW - lower Arctic cost assumptions",
    "BASE": "BASE - central Arctic cost assumptions",
    "HIGH": "HIGH - upper Arctic cost assumptions",
    "MAROCCO_2022": "LITERATURE - Marocco et al. 2022",
    "MCKINLEY_2025": "LITERATURE - McKinley et al. 2025 (Alaska)",
    "JANKE_DEFAULT_2026": "LITERATURE - Janke et al. 2026 (default)",
    "JANKE_SANIRAJAK_2026": "LITERATURE - Janke et al. 2026 (Sanirajak)",
    "MEDIAN": "MEDIAN - median of the 4 literature presets",
}

COST_SCENARIO_DESCRIPTIONS = {
    "LOW": "Lower Arctic/remote sensitivity assumptions from the previous LOW/BASE/HIGH model.",
    "BASE": "Central Arctic/remote sensitivity assumptions from the previous LOW/BASE/HIGH model.",
    "HIGH": "Upper Arctic/remote sensitivity assumptions from the previous LOW/BASE/HIGH model.",
    "MAROCCO_2022": (
        "Marocco et al. 2022 reference economics; electrolyzer and fuel cell use "
        "the published size-dependent cost functions. Existing Hornsund diesel CAPEX is sunk."
    ),
    "MCKINLEY_2025": (
        "Alaska installed-cost sensitivity from McKinley et al. 2025. Uses the "
        "950 kWh BESS row and the EWT wind-cost row; wind is not specific to the 0.6 kW unit."
    ),
    "JANKE_DEFAULT_2026": (
        "Janke et al. 2026 default Arctic-microgrid unit costs. Water treatment, "
        "compressor CAPEX and throughput-based OPEX are explicit."
    ),
    "JANKE_SANIRAJAK_2026": (
        "Janke et al. 2026 Sanirajak community-specific sensitivity. Water treatment, "
        "compressor CAPEX and throughput-based OPEX are explicit."
    ),
    "MEDIAN": (
        "Parameter-by-parameter median of the four named literature presets: "
        "Marocco, McKinley, Janke default and Janke Sanirajak. LOW / BASE / HIGH "
        "are excluded because they are sensitivity scenarios rather than independent "
        "literature sources. EL/FC size scaling is disabled in this synthetic median scenario."
    ),
}

# Every scenario contains exactly the same keys.
COST_SCENARIOS = {
    "LOW": {
        "PV_CAPEX": 2500.0,
        "PV_FIXED_OM": 15.0,
        "PV_VARIABLE_OM": 0.0,
        "PV_REPLACEMENT_FRACTION": 1.0,
        "WIND_CAPEX": 4500.0,
        "WIND_FIXED_OM": 70.0,
        "WIND_VARIABLE_OM": 0.0,
        "WIND_REPLACEMENT_FRACTION": 1.0,
        "BATTERY_CAPEX": 800.0,
        "BATTERY_REPLACEMENT_COST": 500.0,
        "BATTERY_FIXED_OM": 10.0,
        "BATTERY_VARIABLE_OM": 0.0,
        "EL_CAPEX": 4000.0,
        "EL_STACK_REPLACEMENT_FRACTION": 0.2667,
        "EL_TOTAL_OM_FRACTION": 0.04,
        "EL_FIXED_OM_PER_KW_YEAR": 0.0,
        "H2_TANK_CAPEX": 1100.0,
        "H2_TANK_REPLACEMENT_COST": 1100.0,
        "H2_TANK_FIXED_OM_FRACTION": 0.02,
        "H2_TANK_FIXED_OM_PER_KG_YEAR": 0.0,
        "H2_WATER_TREATMENT_CAPEX_PER_KW_EL": 0.0,
        "H2_COMPRESSOR_CAPEX_PER_KW_EL": 0.0,
        "H2_SYSTEM_VARIABLE_OM": 0.0,
        "FC_CAPEX": 4000.0,
        "FC_STACK_REPLACEMENT_FRACTION": 0.2667,
        "FC_TOTAL_OM_FRACTION": 0.04,
        "FC_FIXED_OM_PER_KW_YEAR": 0.0,
        "FC_OM_PER_OPERATING_HOUR": 0.0,
        "DIESEL_CAPEX": 0.0,
        "DIESEL_REPLACEMENT_COST": 420.0,
        "DIESEL_VARIABLE_OM": 0.0,
        "DIESEL_OM_PER_OPERATING_HOUR": 0.40,
        "DIESEL_OM_PER_L": 0.0,
        "DIESEL_FUEL_COST": 1.80,
    },
    "BASE": {
        "PV_CAPEX": 2750.0,
        "PV_FIXED_OM": 17.5,
        "PV_VARIABLE_OM": 0.0,
        "PV_REPLACEMENT_FRACTION": 1.0,
        "WIND_CAPEX": 5000.0,
        "WIND_FIXED_OM": 75.0,
        "WIND_VARIABLE_OM": 0.0,
        "WIND_REPLACEMENT_FRACTION": 1.0,
        "BATTERY_CAPEX": 900.0,
        "BATTERY_REPLACEMENT_COST": 550.0,
        "BATTERY_FIXED_OM": 10.0,
        "BATTERY_VARIABLE_OM": 0.0,
        "EL_CAPEX": 4500.0,
        "EL_STACK_REPLACEMENT_FRACTION": (0.2667 + 0.30) / 2.0,
        "EL_TOTAL_OM_FRACTION": 0.04,
        "EL_FIXED_OM_PER_KW_YEAR": 0.0,
        "H2_TANK_CAPEX": 1225.0,
        "H2_TANK_REPLACEMENT_COST": 1225.0,
        "H2_TANK_FIXED_OM_FRACTION": 0.02,
        "H2_TANK_FIXED_OM_PER_KG_YEAR": 0.0,
        "H2_WATER_TREATMENT_CAPEX_PER_KW_EL": 0.0,
        "H2_COMPRESSOR_CAPEX_PER_KW_EL": 0.0,
        "H2_SYSTEM_VARIABLE_OM": 0.0,
        "FC_CAPEX": 4500.0,
        "FC_STACK_REPLACEMENT_FRACTION": 0.2667,
        "FC_TOTAL_OM_FRACTION": 0.04,
        "FC_FIXED_OM_PER_KW_YEAR": 0.0,
        "FC_OM_PER_OPERATING_HOUR": 0.0,
        "DIESEL_CAPEX": 0.0,
        "DIESEL_REPLACEMENT_COST": 1500.0,
        "DIESEL_VARIABLE_OM": 0.0,
        "DIESEL_OM_PER_OPERATING_HOUR": 0.40,
        "DIESEL_OM_PER_L": 0.0,
        "DIESEL_FUEL_COST": 1.90,
    },
    "HIGH": {
        "PV_CAPEX": 3000.0,
        "PV_FIXED_OM": 20.0,
        "PV_VARIABLE_OM": 0.0,
        "PV_REPLACEMENT_FRACTION": 1.0,
        "WIND_CAPEX": 5500.0,
        "WIND_FIXED_OM": 80.0,
        "WIND_VARIABLE_OM": 0.0,
        "WIND_REPLACEMENT_FRACTION": 1.0,
        "BATTERY_CAPEX": 1000.0,
        "BATTERY_REPLACEMENT_COST": 600.0,
        "BATTERY_FIXED_OM": 10.0,
        "BATTERY_VARIABLE_OM": 0.0,
        "EL_CAPEX": 5000.0,
        "EL_STACK_REPLACEMENT_FRACTION": 0.30,
        "EL_TOTAL_OM_FRACTION": 0.04,
        "EL_FIXED_OM_PER_KW_YEAR": 0.0,
        "H2_TANK_CAPEX": 1350.0,
        "H2_TANK_REPLACEMENT_COST": 1350.0,
        "H2_TANK_FIXED_OM_FRACTION": 0.02,
        "H2_TANK_FIXED_OM_PER_KG_YEAR": 0.0,
        "H2_WATER_TREATMENT_CAPEX_PER_KW_EL": 0.0,
        "H2_COMPRESSOR_CAPEX_PER_KW_EL": 0.0,
        "H2_SYSTEM_VARIABLE_OM": 0.0,
        "FC_CAPEX": 5000.0,
        "FC_STACK_REPLACEMENT_FRACTION": 0.2667,
        "FC_TOTAL_OM_FRACTION": 0.04,
        "FC_FIXED_OM_PER_KW_YEAR": 0.0,
        "FC_OM_PER_OPERATING_HOUR": 0.0,
        "DIESEL_CAPEX": 0.0,
        "DIESEL_REPLACEMENT_COST": 1500.0,
        "DIESEL_VARIABLE_OM": 0.0,
        "DIESEL_OM_PER_OPERATING_HOUR": 0.40,
        "DIESEL_OM_PER_L": 0.0,
        "DIESEL_FUEL_COST": 2.00,
    },
    "MAROCCO_2022": {
        "PV_CAPEX": 1547.0,
        "PV_FIXED_OM": 24.0,
        "PV_VARIABLE_OM": 0.0,
        "PV_REPLACEMENT_FRACTION": 1.0,
        "WIND_CAPEX": 1175.0,
        "WIND_FIXED_OM": 0.03 * 1175.0,
        "WIND_VARIABLE_OM": 0.0,
        "WIND_REPLACEMENT_FRACTION": 1.0,
        "BATTERY_CAPEX": 550.0,
        "BATTERY_REPLACEMENT_COST": 275.0,
        "BATTERY_FIXED_OM": 10.0,
        "BATTERY_VARIABLE_OM": 0.0,
        "EL_CAPEX": 4600.0,
        "EL_STACK_REPLACEMENT_FRACTION": 0.267,
        "EL_TOTAL_OM_FRACTION": 0.04,
        "EL_FIXED_OM_PER_KW_YEAR": 0.0,
        "H2_TANK_CAPEX": 470.0,
        "H2_TANK_REPLACEMENT_COST": 0.0,
        "H2_TANK_FIXED_OM_FRACTION": 0.02,
        "H2_TANK_FIXED_OM_PER_KG_YEAR": 0.0,
        "H2_WATER_TREATMENT_CAPEX_PER_KW_EL": 0.0,
        "H2_COMPRESSOR_CAPEX_PER_KW_EL": 0.0,
        "H2_SYSTEM_VARIABLE_OM": 0.0,
        "FC_CAPEX": 3947.0,
        "FC_STACK_REPLACEMENT_FRACTION": 0.267,
        "FC_TOTAL_OM_FRACTION": 0.04,
        "FC_FIXED_OM_PER_KW_YEAR": 0.0,
        "FC_OM_PER_OPERATING_HOUR": 0.0,
        "DIESEL_CAPEX": 0.0,
        "DIESEL_REPLACEMENT_COST": 420.0,
        "DIESEL_VARIABLE_OM": 0.0,
        "DIESEL_OM_PER_OPERATING_HOUR": 0.40,
        "DIESEL_OM_PER_L": 0.0,
        "DIESEL_FUEL_COST": 2.00,
    },
    "MCKINLEY_2025": {
        "PV_CAPEX": 3250.0 * USD_TO_EUR,
        "PV_FIXED_OM": 10.0 * USD_TO_EUR,
        "PV_VARIABLE_OM": 0.0,
        "PV_REPLACEMENT_FRACTION": 1.0,
        "WIND_CAPEX": 5500.0 * USD_TO_EUR,
        "WIND_FIXED_OM": 75.0 * USD_TO_EUR,
        "WIND_VARIABLE_OM": 0.0,
        "WIND_REPLACEMENT_FRACTION": 1.0,
        "BATTERY_CAPEX": 1000.0 * USD_TO_EUR,
        "BATTERY_REPLACEMENT_COST": 0.75 * 1000.0 * USD_TO_EUR,
        "BATTERY_FIXED_OM": 10.0 * USD_TO_EUR,
        "BATTERY_VARIABLE_OM": 0.0,
        "EL_CAPEX": 2000.0 * USD_TO_EUR,
        "EL_STACK_REPLACEMENT_FRACTION": 0.20,
        "EL_TOTAL_OM_FRACTION": 0.0,
        "EL_FIXED_OM_PER_KW_YEAR": 100.0 * USD_TO_EUR,
        "H2_TANK_CAPEX": 1500.0 * USD_TO_EUR,
        "H2_TANK_REPLACEMENT_COST": 1500.0 * USD_TO_EUR,
        "H2_TANK_FIXED_OM_FRACTION": 0.0,
        "H2_TANK_FIXED_OM_PER_KG_YEAR": 30.0 * USD_TO_EUR,
        "H2_WATER_TREATMENT_CAPEX_PER_KW_EL": 0.0,
        "H2_COMPRESSOR_CAPEX_PER_KW_EL": 0.0,
        "H2_SYSTEM_VARIABLE_OM": 0.0,
        "FC_CAPEX": 2000.0 * USD_TO_EUR,
        "FC_STACK_REPLACEMENT_FRACTION": 0.20,
        "FC_TOTAL_OM_FRACTION": 0.0,
        "FC_FIXED_OM_PER_KW_YEAR": 0.0,
        "FC_OM_PER_OPERATING_HOUR": 0.02 * USD_TO_EUR,
        "DIESEL_CAPEX": 0.0,
        "DIESEL_REPLACEMENT_COST": 420.0,
        "DIESEL_VARIABLE_OM": 0.0,
        "DIESEL_OM_PER_OPERATING_HOUR": 0.40,
        "DIESEL_OM_PER_L": 0.0,
        "DIESEL_FUEL_COST": 2.00,
    },
    "JANKE_DEFAULT_2026": {
        "PV_CAPEX": 2816.0 * CAD_TO_EUR,
        "PV_FIXED_OM": 0.0,
        "PV_VARIABLE_OM": 0.01 * CAD_TO_EUR,
        "PV_REPLACEMENT_FRACTION": 1.0,
        "WIND_CAPEX": 4615.0 * CAD_TO_EUR,
        "WIND_FIXED_OM": 0.0,
        "WIND_VARIABLE_OM": 0.025 * CAD_TO_EUR,
        "WIND_REPLACEMENT_FRACTION": 1.0,
        "BATTERY_CAPEX": 1183.0 * CAD_TO_EUR,
        "BATTERY_REPLACEMENT_COST": 0.50 * 1183.0 * CAD_TO_EUR,
        "BATTERY_FIXED_OM": 0.0,
        "BATTERY_VARIABLE_OM": 0.01 * CAD_TO_EUR,
        "EL_CAPEX": 1600.0 * CAD_TO_EUR,
        "EL_STACK_REPLACEMENT_FRACTION": 0.267,
        "EL_TOTAL_OM_FRACTION": 0.0,
        "EL_FIXED_OM_PER_KW_YEAR": 0.0,
        "H2_TANK_CAPEX": 1200.0 * CAD_TO_EUR,
        "H2_TANK_REPLACEMENT_COST": 0.0,
        "H2_TANK_FIXED_OM_FRACTION": 0.0,
        "H2_TANK_FIXED_OM_PER_KG_YEAR": 0.0,
        "H2_WATER_TREATMENT_CAPEX_PER_KW_EL": 200.0 * CAD_TO_EUR,
        "H2_COMPRESSOR_CAPEX_PER_KW_EL": 2700.0 * CAD_TO_EUR,
        "H2_SYSTEM_VARIABLE_OM": 0.06 * CAD_TO_EUR,
        "FC_CAPEX": 2000.0 * CAD_TO_EUR,
        "FC_STACK_REPLACEMENT_FRACTION": 0.267,
        "FC_TOTAL_OM_FRACTION": 0.0,
        "FC_FIXED_OM_PER_KW_YEAR": 0.0,
        "FC_OM_PER_OPERATING_HOUR": 0.0,
        "DIESEL_CAPEX": 0.0,
        "DIESEL_REPLACEMENT_COST": 420.0,
        "DIESEL_VARIABLE_OM": 0.0,
        "DIESEL_OM_PER_OPERATING_HOUR": 0.40,
        "DIESEL_OM_PER_L": 0.0,
        "DIESEL_FUEL_COST": 2.00,
    },
    "JANKE_SANIRAJAK_2026": {
        "PV_CAPEX": 3617.0 * CAD_TO_EUR,
        "PV_FIXED_OM": 0.0,
        "PV_VARIABLE_OM": 0.01 * CAD_TO_EUR,
        "PV_REPLACEMENT_FRACTION": 1.0,
        "WIND_CAPEX": 10000.0 * CAD_TO_EUR,
        "WIND_FIXED_OM": 0.0,
        "WIND_VARIABLE_OM": 0.025 * CAD_TO_EUR,
        "WIND_REPLACEMENT_FRACTION": 1.0,
        "BATTERY_CAPEX": 2025.0 * CAD_TO_EUR,
        "BATTERY_REPLACEMENT_COST": 0.50 * 2025.0 * CAD_TO_EUR,
        "BATTERY_FIXED_OM": 0.0,
        "BATTERY_VARIABLE_OM": 0.01 * CAD_TO_EUR,
        "EL_CAPEX": 1400.0 * CAD_TO_EUR,
        "EL_STACK_REPLACEMENT_FRACTION": 0.267,
        "EL_TOTAL_OM_FRACTION": 0.0,
        "EL_FIXED_OM_PER_KW_YEAR": 0.0,
        "H2_TANK_CAPEX": 1500.0 * CAD_TO_EUR,
        "H2_TANK_REPLACEMENT_COST": 0.0,
        "H2_TANK_FIXED_OM_FRACTION": 0.0,
        "H2_TANK_FIXED_OM_PER_KG_YEAR": 0.0,
        "H2_WATER_TREATMENT_CAPEX_PER_KW_EL": 200.0 * CAD_TO_EUR,
        "H2_COMPRESSOR_CAPEX_PER_KW_EL": 2700.0 * CAD_TO_EUR,
        "H2_SYSTEM_VARIABLE_OM": 0.06 * CAD_TO_EUR,
        "FC_CAPEX": 3400.0 * CAD_TO_EUR,
        "FC_STACK_REPLACEMENT_FRACTION": 0.267,
        "FC_TOTAL_OM_FRACTION": 0.0,
        "FC_FIXED_OM_PER_KW_YEAR": 0.0,
        "FC_OM_PER_OPERATING_HOUR": 0.0,
        "DIESEL_CAPEX": 0.0,
        "DIESEL_REPLACEMENT_COST": 420.0,
        "DIESEL_VARIABLE_OM": 0.0,
        "DIESEL_OM_PER_OPERATING_HOUR": 0.40,
        "DIESEL_OM_PER_L": 0.0,
        "DIESEL_FUEL_COST": 2.00,
    },
}


# Synthetic eighth scenario: median of each economic input across the four
# named literature presets only. LOW / BASE / HIGH are intentionally excluded
# because they are sensitivity scenarios, not independent literature sources.
# Explicit zeros are included in the median because they are genuine scenario
# inputs (e.g. an O&M channel not used by a source).
_MEDIAN_SOURCE_SCENARIOS = (
    "MAROCCO_2022",
    "MCKINLEY_2025",
    "JANKE_DEFAULT_2026",
    "JANKE_SANIRAJAK_2026",
)
_MEDIAN_COST_KEYS = tuple(
    COST_SCENARIOS[_MEDIAN_SOURCE_SCENARIOS[0]].keys()
)

for _scenario_key in _MEDIAN_SOURCE_SCENARIOS:
    if tuple(COST_SCENARIOS[_scenario_key].keys()) != _MEDIAN_COST_KEYS:
        raise ValueError(
            "All source scenarios must expose the same economic keys before "
            "the MEDIAN scenario can be calculated."
        )

COST_SCENARIOS["MEDIAN"] = {
    _cost_key: float(np.median([
        COST_SCENARIOS[_scenario_key][_cost_key]
        for _scenario_key in _MEDIAN_SOURCE_SCENARIOS
    ]))
    for _cost_key in _MEDIAN_COST_KEYS
}

COST_SCENARIO_META = {
    "LOW": {"el_size_scaling": False, "fc_size_scaling": False},
    "BASE": {"el_size_scaling": False, "fc_size_scaling": False},
    "HIGH": {"el_size_scaling": False, "fc_size_scaling": False},
    "MAROCCO_2022": {"el_size_scaling": True, "fc_size_scaling": True},
    "MCKINLEY_2025": {"el_size_scaling": False, "fc_size_scaling": False},
    "JANKE_DEFAULT_2026": {"el_size_scaling": False, "fc_size_scaling": False},
    "JANKE_SANIRAJAK_2026": {"el_size_scaling": False, "fc_size_scaling": False},
    "MEDIAN": {"el_size_scaling": False, "fc_size_scaling": False},
}

def apply_cost_inputs(costs, scenario_label="CUSTOM", scenario_key=None):
    """Apply one complete set of economic inputs to the simulator globals."""
    global COST_SCENARIO, COST_SCENARIO_KEY
    global PV_CAPEX, PV_FIXED_OM, PV_VARIABLE_OM, PV_REPLACEMENT_FRACTION
    global WIND_CAPEX, WIND_FIXED_OM, WIND_VARIABLE_OM, WIND_REPLACEMENT_FRACTION
    global BATTERY_CAPEX, BATTERY_REPLACEMENT_COST, BATTERY_FIXED_OM, BATTERY_VARIABLE_OM
    global EL_CAPEX, EL_STACK_REPLACEMENT_FRACTION, EL_TOTAL_OM_FRACTION
    global EL_FIXED_OM_PER_KW_YEAR, EL_USE_SIZE_SCALING
    global H2_TANK_CAPEX, H2_TANK_REPLACEMENT_COST, H2_TANK_FIXED_OM_FRACTION
    global H2_TANK_FIXED_OM_PER_KG_YEAR
    global H2_WATER_TREATMENT_CAPEX_PER_KW_EL, H2_COMPRESSOR_CAPEX_PER_KW_EL
    global H2_SYSTEM_VARIABLE_OM
    global FC_CAPEX, FC_STACK_REPLACEMENT_FRACTION, FC_TOTAL_OM_FRACTION
    global FC_FIXED_OM_PER_KW_YEAR, FC_OM_PER_OPERATING_HOUR, FC_USE_SIZE_SCALING
    global DIESEL_CAPEX, DIESEL_REPLACEMENT_COST, DIESEL_VARIABLE_OM
    global DIESEL_OM_PER_OPERATING_HOUR, DIESEL_OM_PER_L, DIESEL_FUEL_COST

    PV_CAPEX = float(costs["PV_CAPEX"])
    PV_FIXED_OM = float(costs["PV_FIXED_OM"])
    PV_VARIABLE_OM = float(costs["PV_VARIABLE_OM"])
    PV_REPLACEMENT_FRACTION = float(costs["PV_REPLACEMENT_FRACTION"])

    WIND_CAPEX = float(costs["WIND_CAPEX"])
    WIND_FIXED_OM = float(costs["WIND_FIXED_OM"])
    WIND_VARIABLE_OM = float(costs["WIND_VARIABLE_OM"])
    WIND_REPLACEMENT_FRACTION = float(costs["WIND_REPLACEMENT_FRACTION"])

    BATTERY_CAPEX = float(costs["BATTERY_CAPEX"])
    BATTERY_REPLACEMENT_COST = float(costs["BATTERY_REPLACEMENT_COST"])
    BATTERY_FIXED_OM = float(costs["BATTERY_FIXED_OM"])
    BATTERY_VARIABLE_OM = float(costs["BATTERY_VARIABLE_OM"])

    EL_CAPEX = float(costs["EL_CAPEX"])
    EL_STACK_REPLACEMENT_FRACTION = float(costs["EL_STACK_REPLACEMENT_FRACTION"])
    EL_TOTAL_OM_FRACTION = float(costs["EL_TOTAL_OM_FRACTION"])
    EL_FIXED_OM_PER_KW_YEAR = float(costs["EL_FIXED_OM_PER_KW_YEAR"])

    H2_TANK_CAPEX = float(costs["H2_TANK_CAPEX"])
    H2_TANK_REPLACEMENT_COST = float(costs["H2_TANK_REPLACEMENT_COST"])
    H2_TANK_FIXED_OM_FRACTION = float(costs["H2_TANK_FIXED_OM_FRACTION"])
    H2_TANK_FIXED_OM_PER_KG_YEAR = float(costs["H2_TANK_FIXED_OM_PER_KG_YEAR"])
    H2_WATER_TREATMENT_CAPEX_PER_KW_EL = float(
        costs["H2_WATER_TREATMENT_CAPEX_PER_KW_EL"]
    )
    H2_COMPRESSOR_CAPEX_PER_KW_EL = float(
        costs["H2_COMPRESSOR_CAPEX_PER_KW_EL"]
    )
    H2_SYSTEM_VARIABLE_OM = float(costs["H2_SYSTEM_VARIABLE_OM"])

    FC_CAPEX = float(costs["FC_CAPEX"])
    FC_STACK_REPLACEMENT_FRACTION = float(costs["FC_STACK_REPLACEMENT_FRACTION"])
    FC_TOTAL_OM_FRACTION = float(costs["FC_TOTAL_OM_FRACTION"])
    FC_FIXED_OM_PER_KW_YEAR = float(costs["FC_FIXED_OM_PER_KW_YEAR"])
    FC_OM_PER_OPERATING_HOUR = float(costs["FC_OM_PER_OPERATING_HOUR"])

    DIESEL_CAPEX = float(costs["DIESEL_CAPEX"])
    DIESEL_REPLACEMENT_COST = float(costs["DIESEL_REPLACEMENT_COST"])
    DIESEL_VARIABLE_OM = float(costs["DIESEL_VARIABLE_OM"])
    DIESEL_OM_PER_OPERATING_HOUR = float(costs["DIESEL_OM_PER_OPERATING_HOUR"])
    DIESEL_OM_PER_L = float(costs["DIESEL_OM_PER_L"])
    DIESEL_FUEL_COST = float(costs["DIESEL_FUEL_COST"])

    if scenario_key is not None and scenario_key in COST_SCENARIO_META:
        COST_SCENARIO_KEY = scenario_key
        EL_USE_SIZE_SCALING = bool(COST_SCENARIO_META[scenario_key]["el_size_scaling"])
        FC_USE_SIZE_SCALING = bool(COST_SCENARIO_META[scenario_key]["fc_size_scaling"])

    COST_SCENARIO = str(scenario_label)


def apply_cost_scenario(name):
    """Apply one of the source-based economic scenarios."""
    raw = str(name)
    key = raw.upper()
    if key not in COST_SCENARIOS:
        reverse_labels = {
            label.upper(): scenario_key
            for scenario_key, label in COST_SCENARIO_LABELS.items()
        }
        key = reverse_labels.get(raw.upper(), key)
    if key not in COST_SCENARIOS:
        raise ValueError(
            "Unknown cost scenario. Available: " + ", ".join(COST_SCENARIOS)
        )
    apply_cost_inputs(
        COST_SCENARIOS[key],
        scenario_label=COST_SCENARIO_LABELS.get(key, key),
        scenario_key=key,
    )


def electrolyzer_specific_capex(rated_power_kw=None):
    """Return EL specific CAPEX [EUR/kW] for the active economic scenario."""
    rated = EL_SIZE if rated_power_kw is None else float(rated_power_kw)
    if rated <= 0:
        return 0.0
    if EL_USE_SIZE_SCALING:
        return float(
            EL_CAPEX
            * (rated / EL_CAPEX_REFERENCE_SIZE) ** (EL_COST_EXPONENT - 1.0)
        )
    return float(EL_CAPEX)


def fuel_cell_specific_capex(rated_power_kw=None):
    """Return FC specific CAPEX [EUR/kW] for the active economic scenario."""
    rated = FC_SIZE if rated_power_kw is None else float(rated_power_kw)
    if rated <= 0:
        return 0.0
    if FC_USE_SIZE_SCALING:
        return float(
            FC_CAPEX
            * (rated / FC_CAPEX_REFERENCE_SIZE) ** (FC_COST_EXPONENT - 1.0)
        )
    return float(FC_CAPEX)


def h2_auxiliary_capex():
    """Water-treatment + compressor CAPEX tied to EL rated power [EUR]."""
    return float(
        (H2_WATER_TREATMENT_CAPEX_PER_KW_EL + H2_COMPRESSOR_CAPEX_PER_KW_EL)
        * EL_SIZE
    )


# BASE (central Arctic sensitivity) is the default for direct runs and optimizer imports.
apply_cost_scenario(COST_SCENARIO)

# --- Alternatore (Mecc Alte ECO 34-1SN/4, 400 V / 50 Hz, classe H) -----------
# DIESEL_SIZE, dg_min, dg_max e il dispatch restano in kW ELETTRICI al quadro.
# Il rendimento dell'alternatore serve solo a risalire alla potenza all'albero
# su cui e' definita la curva di consumo:  P_albero = P_el / eta_alt(carico).
# Il carico in per unita' dei grafici del datasheet e' riferito ai kVA nominali:
#   carico_pu = P_el / (cos_phi * kVA_nominali)
ALTERNATOR_EFFICIENCY_ENABLED = True
ALTERNATOR_RATED_KVA = 85.0          # kVA, classe H, 400 V / 50 Hz
ALTERNATOR_POWER_FACTOR = 0.8        # cos phi del carico (nel modello non c'e' potenza reattiva)
ALT_LOAD_PU = np.array([0.25, 0.50, 0.75, 1.00])        # carico [p.u.]
ALT_EFF_PF08 = np.array([86.0, 90.0, 91.9, 91.5])       # %, da tabella datasheet (cos phi = 0,8)
ALT_EFF_PF10 = np.array([88.0, 92.0, 93.8, 93.4])       # %, dal grafico (cos phi = 1)


# High-contrast, colour-blind-friendly plotting palette.
COLORS = {
    "load": "#202020", "pv": "#E69F00", "wind": "#0072B2",
    "renewables": "#009E73", "battery": "#CC79A7", "electrolyzer": "#56B4E9",
    "fuel_cell": "#6A3D9A", "diesel": "#D55E00", "curtailment": "#999999",
    "unserved": "#E41A1C", "h2": "#00A6D6",
}

COST_COLORS = {
    "Wind": "#4C72B0",
    "PV": "#DDAB41",
    "Battery Li-ion": "#8172B2",
    "PEM electrolyzer": "#64B5CD",
    "PEM fuel cell": "#55A868",
    "H2 tank": "#8DD3C7",
    "H2 auxiliaries": "#B3DE69",
    "Diesel generator": "#C44E52",
}


def alternator_efficiency(p_el_kw):
    """Rendimento dell'alternatore [-] alla potenza attiva erogata [kW].

    Interpola le curve del datasheet a cos phi 0,8 e 1 in funzione del carico
    in p.u. (kVA/kVA nominali) e le combina linearmente sul cos phi impostato.
    Fuori dal campo dei grafici (< 0,25 p.u.) il valore viene tenuto costante.
    """
    if not ALTERNATOR_EFFICIENCY_ENABLED:
        return 1.0
    load_pu = p_el_kw / (ALTERNATOR_RATED_KVA * ALTERNATOR_POWER_FACTOR)
    eta_08 = np.interp(load_pu, ALT_LOAD_PU, ALT_EFF_PF08)
    eta_10 = np.interp(load_pu, ALT_LOAD_PU, ALT_EFF_PF10)
    weight = np.clip((ALTERNATOR_POWER_FACTOR - 0.8) / 0.2, 0.0, 1.0)
    return float((eta_08 + weight * (eta_10 - eta_08)) / 100)


def electrolyzer_efficiency(p_el_kw, rated_power_kw=None):
    """PEM electrolyzer system efficiency [-], LHV basis, at actual load.

    The curve is used only inside the admissible modulation range. Below the
    technical minimum the electrolyzer is considered OFF and efficiency is 0.
    """
    rated = EL_SIZE if rated_power_kw is None else float(rated_power_kw)
    p_el_kw = float(p_el_kw)
    if rated <= 0 or p_el_kw <= 0:
        return 0.0
    load_fraction = p_el_kw / rated
    if load_fraction < EL_MIN_LOAD_FRACTION - 1e-12:
        return 0.0
    load_fraction = float(np.clip(load_fraction, EL_MIN_LOAD_FRACTION, 1.0))
    return float(np.interp(
        load_fraction, EL_EFFICIENCY_LOAD_FRACTION, EL_EFFICIENCY_LHV
    ))


def fuel_cell_efficiency(p_fc_kw, rated_power_kw=None):
    """PEM fuel-cell system efficiency [-], LHV basis, at actual load.

    Below the technical minimum the fuel cell is considered OFF.
    """
    rated = FC_SIZE if rated_power_kw is None else float(rated_power_kw)
    p_fc_kw = float(p_fc_kw)
    if rated <= 0 or p_fc_kw <= 0:
        return 0.0
    load_fraction = p_fc_kw / rated
    if load_fraction < FC_MIN_LOAD_FRACTION - 1e-12:
        return 0.0
    load_fraction = float(np.clip(load_fraction, FC_MIN_LOAD_FRACTION, 1.0))
    return float(np.interp(
        load_fraction, FC_EFFICIENCY_LOAD_FRACTION, FC_EFFICIENCY_LHV
    ))


def max_electrolyzer_power_for_h2_room(h2_room_kwh, rated_power_kw=None):
    """Max EL input power [kW] that fits the available H2-energy room.

    The hourly model has dt = 1 h, so produced H2 energy is
    P_EL * eta_EL(P_EL). If the H2 room is too small even for the minimum
    technical load, the electrolyzer must remain OFF.
    """
    rated = EL_SIZE if rated_power_kw is None else float(rated_power_kw)
    room = max(0.0, float(h2_room_kwh))
    if rated <= 0 or room <= 0:
        return 0.0

    p_min = EL_MIN_LOAD_FRACTION * rated
    eta_min = electrolyzer_efficiency(p_min, rated)
    if eta_min <= 0 or room + 1e-12 < p_min * eta_min:
        return 0.0

    eta_max = electrolyzer_efficiency(rated, rated)
    if rated * eta_max <= room + 1e-12:
        return rated

    lo, hi = p_min, rated
    for _ in range(60):
        mid = 0.5 * (lo + hi)
        eta_mid = electrolyzer_efficiency(mid, rated)
        if mid * eta_mid <= room:
            lo = mid
        else:
            hi = mid
    return lo


def max_fuel_cell_power_from_h2(h2_available_kwh, rated_power_kw=None):
    """Max FC electrical output [kW] allowed by available H2 energy.

    The hourly H2 consumption is P_FC / eta_FC(P_FC). If the stored H2
    cannot sustain the minimum technical power for one hour, the FC is OFF.
    """
    rated = FC_SIZE if rated_power_kw is None else float(rated_power_kw)
    available = max(0.0, float(h2_available_kwh))
    if rated <= 0 or available <= 0:
        return 0.0

    p_min = FC_MIN_LOAD_FRACTION * rated
    eta_min = fuel_cell_efficiency(p_min, rated)
    if eta_min <= 0 or available + 1e-12 < p_min / eta_min:
        return 0.0

    eta_max = fuel_cell_efficiency(rated, rated)
    if rated / eta_max <= available + 1e-12:
        return rated

    lo, hi = p_min, rated
    for _ in range(60):
        mid = 0.5 * (lo + hi)
        eta_mid = fuel_cell_efficiency(mid, rated)
        if mid / eta_mid <= available:
            lo = mid
        else:
            hi = mid
    return lo


def _component_is_active(component):
    """True se il componente e' attivo nella configurazione selezionata."""
    active = {
        "Battery Li-ion": BATTERY_SIZE > 0,
        "PEM electrolyzer": EL_SIZE > 0,
        "PEM fuel cell": FC_SIZE > 0,
        "H2 storage unit": H2_STORAGE_SIZE > 0,
        "Wind turbine unit": N_WT > 0 and WIND_SIZE > 0,
        "PV shipment item": PV_SIZE > 0,
    }
    return active.get(component, True)


def calculate_design_constraints():
    """Valuta i vincoli Hornsund per la configurazione attiva.

    Sono check di fattibilita': non ridimensionano automaticamente i componenti.
    """
    h2_mass_kg = H2_STORAGE_SIZE * 3.6 / H2_LHV if H2_STORAGE_SIZE > 0 else 0.0
    h2_system_capex = (
        electrolyzer_specific_capex(EL_SIZE) * EL_SIZE
        + fuel_cell_specific_capex(FC_SIZE) * FC_SIZE
        + H2_TANK_CAPEX * h2_mass_kg
        + h2_auxiliary_capex()
    )

    pv_size_feasible = PV_SIZE <= PV_MAX_SIZE + 1e-12
    wind_size_feasible = N_WT <= N_WT_MAX
    # Nei casi Hornsund con batteria, la capacita' prevista e' trattata come
    # taglia fissa di progetto (500 kWh), non come semplice limite superiore.
    # Il caso diesel-only non utilizza la batteria e quindi soddisfa
    # automaticamente questo check.
    if SYSTEM_MODE == "diesel_only":
        battery_size_feasible = True
    else:
        battery_size_feasible = (
            abs(BATTERY_SIZE - BATTERY_HORNSUND_FIXED_SIZE) <= 1e-9
        )

    h2_budget_feasible = h2_system_capex <= H2_SYSTEM_BUDGET + 1e-12

    item_mass_checks = {}
    for component, mass_kg in COMPONENT_ITEM_MASS_KG.items():
        if not _component_is_active(component):
            item_mass_checks[component] = {"active": False}
        elif mass_kg is None:
            item_mass_checks[component] = None
        else:
            item_mass_checks[component] = {
                "active": True,
                "mass_kg": float(mass_kg),
                "margin_kg": MAX_MASS_PER_CRANE_ITEM_KG - float(mass_kg),
                "feasible": float(mass_kg) <= MAX_MASS_PER_CRANE_ITEM_KG,
            }

    known_item_masses = [
        float(mass_kg)
        for component, mass_kg in COMPONENT_ITEM_MASS_KG.items()
        if mass_kg is not None and _component_is_active(component)
    ]
    known_total_mass_kg = sum(known_item_masses)

    known_mass_checks = [
        check["feasible"]
        for check in item_mass_checks.values()
        if isinstance(check, dict)
        and check.get("active", False)
        and "feasible" in check
    ]
    crane_feasible = all(known_mass_checks) if known_mass_checks else True

    overall_pre_feasible = (
        pv_size_feasible
        and wind_size_feasible
        and battery_size_feasible
        and h2_budget_feasible
        and crane_feasible
    )

    return {
        "pv_size_feasible": pv_size_feasible,
        "pv_size_margin_kw": PV_MAX_SIZE - PV_SIZE,
        "wind_size_feasible": wind_size_feasible,
        "wind_size_margin_units": N_WT_MAX - N_WT,
        "battery_size_feasible": battery_size_feasible,
        "battery_size_margin_kwh": BATTERY_HORNSUND_FIXED_SIZE - BATTERY_SIZE,
        "h2_system_capex": h2_system_capex,
        "h2_budget_margin": H2_SYSTEM_BUDGET - h2_system_capex,
        "h2_budget_feasible": h2_budget_feasible,
        "item_mass_checks": item_mass_checks,
        "known_total_mass_kg": known_total_mass_kg,
        "single_shipment_margin_kg": (
            MAX_MASS_PER_SHIPMENT_KG - known_total_mass_kg
            if known_item_masses else None
        ),
        "overall_pre_feasible": overall_pre_feasible,
    }


def print_design_constraint_check(constraints):
    """Stampa i check dei vincoli Hornsund."""
    print("\nHORNSUND DESIGN CONSTRAINTS: ENABLED")

    status = "OK" if constraints["pv_size_feasible"] else "NOT FEASIBLE"
    print(
        f"{'PV installed power:':35s}{PV_SIZE:12.2f} kW "
        f"(max {PV_MAX_SIZE:.2f} kW) [{status}]"
    )

    status = "OK" if constraints["wind_size_feasible"] else "NOT FEASIBLE"
    print(
        f"{'Number of wind turbines:':35s}{N_WT:12d} "
        f"(max {N_WT_MAX:d}) [{status}]"
    )

    status = "OK" if constraints["battery_size_feasible"] else "NOT FEASIBLE"
    if SYSTEM_MODE == "diesel_only":
        print(f"{'Battery capacity:':35s}{'NOT ACTIVE':>12s} [OK]")
    else:
        print(
            f"{'Battery capacity:':35s}{BATTERY_SIZE:12.2f} kWh "
            f"(fixed Hornsund size {BATTERY_HORNSUND_FIXED_SIZE:.2f} kWh) [{status}]"
        )

    h2_capex = constraints["h2_system_capex"]
    h2_margin = constraints["h2_budget_margin"]
    status = "OK" if constraints["h2_budget_feasible"] else "NOT FEASIBLE"
    print(
        f"{'H2-system CAPEX:':35s}{h2_capex:12.2f} EUR "
        f"(budget {H2_SYSTEM_BUDGET:.2f} EUR) [{status}]"
    )
    print(f"{'H2 budget margin:':35s}{h2_margin:12.2f} EUR")

    print(
        f"{'Max mass per crane item:':35s}"
        f"{MAX_MASS_PER_CRANE_ITEM_KG:12.1f} kg"
    )
    for component, check in constraints["item_mass_checks"].items():
        if isinstance(check, dict) and not check.get("active", True):
            print(f"  {component + ':':33s}NOT ACTIVE")
        elif check is None:
            print(f"  {component + ':':33s}TO BE DEFINED")
        else:
            status = "OK" if check["feasible"] else "NOT FEASIBLE"
            print(
                f"  {component + ':':33s}{check['mass_kg']:10.1f} kg [{status}]"
            )

    print(
        f"{'Max mass per sea shipment:':35s}"
        f"{MAX_MASS_PER_SHIPMENT_KG:12.1f} kg"
    )
    print(
        "  Note: the 10 t value is informational at system level because "
        "multiple sea shipments may be possible."
    )

    overall = "OK" if constraints["overall_pre_feasible"] else "NOT FEASIBLE"
    print(f"{'Pre-simulation constraint status:':35s}{overall}")


def print_fuel_constraint_check(r):
    """Controlla il limite annuale di gasolio dopo la simulazione."""
    fuel = float(r["diesel_fuel"].sum())
    feasible = fuel <= MAX_DIESEL_FUEL_L_PER_YEAR + 1e-12
    status = "OK" if feasible else "NOT FEASIBLE"
    print("\nPOST-SIMULATION CONSTRAINT CHECK")
    print(
        f"{'Annual diesel fuel:':35s}{fuel:12.2f} L "
        f"(max {MAX_DIESEL_FUEL_L_PER_YEAR:.2f} L/year) [{status}]"
    )
    return feasible


def check_battery_terminal_soc(r, print_result=True):
    """Check SOC_battery,end >= SOC_battery,initial."""
    if BATTERY_SIZE <= 0:
        if print_result:
            print("\nBATTERY TERMINAL SOC CHECK")
            print(f"{'Battery:':35s}{'NOT ACTIVE':>12s} [OK]")
        return True

    soc_initial = float(r.attrs.get("battery_soc_initial", BATTERY_SOC_INITIAL))
    soc_final = float(r.attrs["battery_soc_final"])
    difference = soc_final - soc_initial
    feasible = difference >= -BATTERY_SOC_END_TOLERANCE - 1e-12

    if print_result:
        status = "OK" if feasible else "NOT FEASIBLE"
        print("\nBATTERY TERMINAL SOC CHECK")
        print(f"{'Initial battery SOC:':35s}{100 * soc_initial:12.3f} %")
        print(f"{'Final battery SOC:':35s}{100 * soc_final:12.3f} %")
        print(f"{'SOC end - SOC initial:':35s}{100 * difference:12.3f} %-points")
        print(
            f"{'Terminal SOC condition:':35s}"
            f"SOC_end >= SOC_initial [{status}]"
        )
    return feasible


def check_h2_terminal_loh(r, print_result=True):
    """Check LOH_H2,end >= LOH_H2,initial."""
    if H2_STORAGE_SIZE <= 0:
        if print_result:
            print("\nH2 TERMINAL LOH CHECK")
            print(f"{'H2 storage:':35s}{'NOT ACTIVE':>12s} [OK]")
        return True

    loh_initial = float(
        r.attrs.get("h2_loh_initial", r.attrs.get("h2_soc_initial", H2_SOC_INITIAL))
    )
    loh_final = float(
        r.attrs.get("h2_loh_final", r.attrs["h2_soc_final"])
    )
    difference = loh_final - loh_initial
    feasible = difference >= -H2_LOH_END_TOLERANCE - 1e-12

    if print_result:
        status = "OK" if feasible else "NOT FEASIBLE"
        print("\nH2 TERMINAL LOH CHECK")
        print(f"{'Initial H2 LOH:':35s}{100 * loh_initial:12.3f} %")
        print(f"{'Final H2 LOH:':35s}{100 * loh_final:12.3f} %")
        print(f"{'LOH end - LOH initial:':35s}{100 * difference:12.3f} %-points")
        print(
            f"{'Terminal LOH condition:':35s}"
            f"LOH_end >= LOH_initial [{status}]"
        )
    return feasible

def select_simulation_options():
    """Chiede vincoli ON/OFF e il caso energetico da simulare, inclusi i sottocasi battery-only."""
    root = tk.Tk()
    root.title("Off-grid simulation setup")
    root.resizable(False, False)
    root.attributes("-topmost", True)

    apply_constraints_var = tk.BooleanVar(value=True)
    mode_var = tk.StringVar(value="mix")
    accepted = {"value": False}

    frame = tk.Frame(root, padx=18, pady=16)
    frame.pack(fill="both", expand=True)

    tk.Label(
        frame,
        text="Simulation options",
        font=("Arial", 12, "bold")
    ).pack(anchor="w", pady=(0, 10))

    tk.Checkbutton(
        frame,
        text="Apply Hornsund design constraints",
        variable=apply_constraints_var
    ).pack(anchor="w", pady=(0, 12))

    tk.Label(
        frame,
        text="Choose the energy-system configuration:"
    ).pack(anchor="w")

    for label, value in [
        ("Diesel only", "diesel_only"),
        ("Hybrid mix — Battery + H2 + diesel", "mix"),
        ("100% renewable — Battery + H2 (no diesel)", "renewable_only"),
        ("Hybrid mix — Battery only + diesel", "battery_only_mix"),
        ("100% renewable — Battery only (no diesel)", "battery_only_renewable"),
    ]:
        tk.Radiobutton(
            frame, text=label, variable=mode_var, value=value
        ).pack(anchor="w", padx=(12, 0))

    button_frame = tk.Frame(frame)
    button_frame.pack(fill="x", pady=(16, 0))

    def accept():
        accepted["value"] = True
        root.destroy()

    def cancel():
        root.destroy()

    tk.Button(button_frame, text="Start", width=12, command=accept).pack(
        side="right", padx=(6, 0)
    )
    tk.Button(button_frame, text="Cancel", width=12, command=cancel).pack(
        side="right"
    )

    root.protocol("WM_DELETE_WINDOW", cancel)
    root.mainloop()

    if not accepted["value"]:
        raise SystemExit("Simulation cancelled by user.")

    return bool(apply_constraints_var.get()), mode_var.get()



def select_component_sizes(mode):
    """First GUI screen: edit only the component sizes for the selected mode."""
    global PV_SIZE, N_WT, WIND_SIZE, BATTERY_SIZE
    global EL_SIZE, H2_STORAGE_SIZE, FC_SIZE, DIESEL_SIZE

    root = tk.Tk()
    root.title("1/2 - Component sizes")
    root.resizable(False, False)
    root.attributes("-topmost", True)

    accepted = {"value": False}

    pv_var = tk.StringVar(value=f"{PV_SIZE:g}")
    n_wt_var = tk.StringVar(value=f"{N_WT:d}")
    battery_var = tk.StringVar(value=f"{BATTERY_SIZE:g}")
    el_var = tk.StringVar(value=f"{EL_SIZE:g}")
    h2_var = tk.StringVar(value=f"{H2_STORAGE_SIZE:g}")
    fc_var = tk.StringVar(value=f"{FC_SIZE:g}")
    diesel_var = tk.StringVar(value=f"{DIESEL_SIZE:g}")
    wind_total_var = tk.StringVar()
    h2_mass_var = tk.StringVar()

    main = tk.Frame(root, padx=16, pady=14)
    main.pack(fill="both", expand=True)

    tk.Label(
        main,
        text="Component sizes",
        font=("Arial", 12, "bold"),
    ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 4))

    tk.Label(
        main,
        text=(
            "Modify the component sizes for this simulation. Components that are not "
            "part of the selected architecture are disabled automatically."
        ),
        justify="left",
        fg="gray35",
        wraplength=500,
    ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 12))

    def add_row(row, label, variable, unit):
        tk.Label(main, text=label + ":").grid(
            row=row, column=0, sticky="e", padx=(0, 8), pady=4
        )
        entry = tk.Entry(main, textvariable=variable, width=15)
        entry.grid(row=row, column=1, sticky="w", pady=4)
        tk.Label(main, text=unit).grid(
            row=row, column=2, sticky="w", padx=(6, 0), pady=4
        )
        return entry

    pv_entry = add_row(2, "PV installed power", pv_var, "kW")
    wind_entry = add_row(3, "Number of wind turbines", n_wt_var, "units")
    battery_entry = add_row(4, "Battery capacity", battery_var, "kWh")
    el_entry = add_row(5, "Electrolyzer rated power", el_var, "kW")
    h2_entry = add_row(6, "Hydrogen storage", h2_var, "kWh_H2 LHV")
    fc_entry = add_row(7, "Fuel cell rated power", fc_var, "kW")
    diesel_entry = add_row(8, "Diesel rated power", diesel_var, "kW")

    tk.Label(main, textvariable=wind_total_var, fg="gray35").grid(
        row=9, column=0, columnspan=3, sticky="w", pady=(10, 0)
    )
    tk.Label(main, textvariable=h2_mass_var, fg="gray35").grid(
        row=10, column=0, columnspan=3, sticky="w", pady=(2, 0)
    )
    tk.Label(
        main,
        text=(
            f"Wind model: {WT_RATED_POWER:.3f} kW per turbine. The rated power is kept "
            "fixed because the wind capacity-factor input refers to this turbine model."
        ),
        justify="left",
        fg="gray35",
        wraplength=500,
    ).grid(row=11, column=0, columnspan=3, sticky="w", pady=(8, 0))

    def update_previews(*_):
        try:
            n_wt = int(n_wt_var.get())
            wind_total_var.set(
                f"Installed wind power = {n_wt} x {WT_RATED_POWER:.3f} "
                f"= {n_wt * WT_RATED_POWER:.2f} kW"
            )
        except ValueError:
            wind_total_var.set("Installed wind power: enter an integer number of turbines.")

        try:
            h2_kwh = float(h2_var.get())
            h2_kg = h2_kwh * 3.6 / H2_LHV if h2_kwh >= 0 else float("nan")
            h2_mass_var.set(
                f"Equivalent hydrogen mass = {h2_kg:.2f} kg H2"
                if np.isfinite(h2_kg) else ""
            )
        except ValueError:
            h2_mass_var.set("Hydrogen mass: enter a numeric storage value.")

    n_wt_var.trace_add("write", update_previews)
    h2_var.trace_add("write", update_previews)
    update_previews()

    # Disable sizes excluded by the selected architecture.
    if mode == "diesel_only":
        for entry in (pv_entry, wind_entry, battery_entry, el_entry, h2_entry, fc_entry):
            entry.configure(state="disabled")
    elif mode == "renewable_only":
        diesel_entry.configure(state="disabled")
    elif mode == "battery_only_mix":
        for entry in (el_entry, h2_entry, fc_entry):
            entry.configure(state="disabled")
    elif mode == "battery_only_renewable":
        for entry in (el_entry, h2_entry, fc_entry, diesel_entry):
            entry.configure(state="disabled")

    button_frame = tk.Frame(main)
    button_frame.grid(row=12, column=0, columnspan=3, sticky="e", pady=(14, 0))

    def accept():
        try:
            sizes = {
                "pv": float(pv_var.get()),
                "n_wt": int(n_wt_var.get()),
                "battery": float(battery_var.get()),
                "el": float(el_var.get()),
                "h2": float(h2_var.get()),
                "fc": float(fc_var.get()),
                "diesel": float(diesel_var.get()),
            }
        except ValueError:
            messagebox.showerror(
                "Invalid input",
                "All sizes must be numeric and the number of wind turbines must be an integer.",
                parent=root,
            )
            return

        if any(value < 0 for value in sizes.values()):
            messagebox.showerror(
                "Invalid input",
                "Component sizes cannot be negative.",
                parent=root,
            )
            return

        accepted["value"] = True
        accepted["sizes"] = sizes
        root.destroy()

    def cancel():
        root.destroy()

    tk.Button(button_frame, text="Continue to costs", width=16, command=accept).pack(
        side="right", padx=(6, 0)
    )
    tk.Button(button_frame, text="Cancel", width=12, command=cancel).pack(side="right")

    root.protocol("WM_DELETE_WINDOW", cancel)
    root.mainloop()

    if not accepted["value"]:
        raise SystemExit("Simulation cancelled by user.")

    sizes = accepted["sizes"]
    PV_SIZE = sizes["pv"]
    N_WT = sizes["n_wt"]
    WIND_SIZE = N_WT * WT_RATED_POWER
    BATTERY_SIZE = sizes["battery"]
    EL_SIZE = sizes["el"]
    H2_STORAGE_SIZE = sizes["h2"]
    FC_SIZE = sizes["fc"]
    DIESEL_SIZE = sizes["diesel"]

    # Selected architecture always has priority over manual size fields.
    if mode == "diesel_only":
        PV_SIZE = 0.0
        N_WT = 0
        WIND_SIZE = 0.0
        BATTERY_SIZE = 0.0
        EL_SIZE = 0.0
        H2_STORAGE_SIZE = 0.0
        FC_SIZE = 0.0
    elif mode == "renewable_only":
        DIESEL_SIZE = 0.0
    elif mode == "battery_only_mix":
        EL_SIZE = 0.0
        H2_STORAGE_SIZE = 0.0
        FC_SIZE = 0.0
    elif mode == "battery_only_renewable":
        EL_SIZE = 0.0
        H2_STORAGE_SIZE = 0.0
        FC_SIZE = 0.0
        DIESEL_SIZE = 0.0

    h2_mass_kg = H2_STORAGE_SIZE * 3.6 / H2_LHV if H2_STORAGE_SIZE > 0 else 0.0

    print("\nCOMPONENT SIZES USED FOR THIS RUN")
    print(f"  PV:              {PV_SIZE:.2f} kW")
    print(f"  Wind turbines:   {N_WT:d} x {WT_RATED_POWER:.3f} kW = {WIND_SIZE:.2f} kW")
    print(f"  Battery:         {BATTERY_SIZE:.2f} kWh")
    print(f"  Electrolyzer:    {EL_SIZE:.2f} kW")
    print(f"  H2 storage:      {H2_STORAGE_SIZE:.2f} kWh_H2 ({h2_mass_kg:.2f} kg H2)")
    print(f"  Fuel cell:       {FC_SIZE:.2f} kW")
    print(f"  Diesel:          {DIESEL_SIZE:.2f} kW")


def select_economic_inputs():
    """Second GUI screen: choose one of 7 cost scenarios and optionally edit it."""
    global COST_SCENARIO

    root = tk.Tk()
    root.title("2/2 - Economic inputs")
    root.attributes("-topmost", True)

    # Keep the economic-input window usable also on laptops / Windows scaling.
    # The content is vertically scrollable and the action buttons stay fixed
    # at the bottom of the window.
    screen_w = root.winfo_screenwidth()
    screen_h = root.winfo_screenheight()
    window_w = min(1000, max(760, screen_w - 100))
    window_h = min(820, max(520, screen_h - 160))
    root.geometry(f"{window_w}x{window_h}")
    root.minsize(720, 480)
    root.resizable(True, True)

    accepted = {"value": False}
    initial_scenario = (
        COST_SCENARIO_KEY if COST_SCENARIO_KEY in COST_SCENARIOS else "BASE"
    )
    scenario_var = tk.StringVar(value=initial_scenario)
    description_var = tk.StringVar()

    # Every scenario uses the same keys. Separate fixed and throughput-based OPEX
    # are shown explicitly because Marocco, McKinley and Janke use different bases.
    groups = [
        ("PV", [
            ("PV_CAPEX", "CAPEX", "EUR/kW", 1.0),
            ("PV_FIXED_OM", "Fixed O&M", "EUR/(kW y)", 1.0),
            ("PV_VARIABLE_OM", "Variable O&M", "EUR/kWh gen.", 1.0),
            ("PV_REPLACEMENT_FRACTION", "Replacement", "% CAPEX", 100.0),
        ]),
        ("Wind", [
            ("WIND_CAPEX", "CAPEX", "EUR/kW", 1.0),
            ("WIND_FIXED_OM", "Fixed O&M", "EUR/(kW y)", 1.0),
            ("WIND_VARIABLE_OM", "Variable O&M", "EUR/kWh gen.", 1.0),
            ("WIND_REPLACEMENT_FRACTION", "Replacement", "% CAPEX", 100.0),
        ]),
        ("Battery Li-ion", [
            ("BATTERY_CAPEX", "CAPEX", "EUR/kWh", 1.0),
            ("BATTERY_REPLACEMENT_COST", "Replacement", "EUR/kWh", 1.0),
            ("BATTERY_FIXED_OM", "Fixed O&M", "EUR/(kWh y)", 1.0),
            ("BATTERY_VARIABLE_OM", "Variable O&M", "EUR/kWh throughput", 1.0),
        ]),
        ("PEM electrolyzer", [
            ("EL_CAPEX", "CAPEX / ref. CAPEX", "EUR/kW", 1.0),
            ("EL_STACK_REPLACEMENT_FRACTION", "Stack replacement", "% CAPEX", 100.0),
            ("EL_TOTAL_OM_FRACTION", "Marocco total O&M", "% CAPEX/y", 100.0),
            ("EL_FIXED_OM_PER_KW_YEAR", "Fixed O&M", "EUR/(kW y)", 1.0),
        ]),
        ("H2 tank + auxiliaries", [
            ("H2_TANK_CAPEX", "Tank CAPEX", "EUR/kg H2", 1.0),
            ("H2_TANK_REPLACEMENT_COST", "Tank replacement", "EUR/kg H2", 1.0),
            ("H2_TANK_FIXED_OM_FRACTION", "Tank fixed O&M", "% CAPEX/y", 100.0),
            ("H2_TANK_FIXED_OM_PER_KG_YEAR", "Tank fixed O&M", "EUR/(kg y)", 1.0),
            ("H2_WATER_TREATMENT_CAPEX_PER_KW_EL", "Water treatment CAPEX", "EUR/kW_EL", 1.0),
            ("H2_COMPRESSOR_CAPEX_PER_KW_EL", "Compressor CAPEX", "EUR/kW_EL", 1.0),
            ("H2_SYSTEM_VARIABLE_OM", "H2-system O&M", "EUR/kWh H2 thr.", 1.0),
        ]),
        ("PEM fuel cell", [
            ("FC_CAPEX", "CAPEX / ref. CAPEX", "EUR/kW", 1.0),
            ("FC_STACK_REPLACEMENT_FRACTION", "Stack replacement", "% CAPEX", 100.0),
            ("FC_TOTAL_OM_FRACTION", "Marocco total O&M", "% CAPEX/y", 100.0),
            ("FC_FIXED_OM_PER_KW_YEAR", "Fixed O&M", "EUR/(kW y)", 1.0),
            ("FC_OM_PER_OPERATING_HOUR", "Operating O&M", "EUR/op. h", 1.0),
        ]),
        ("Diesel generator", [
            ("DIESEL_CAPEX", "Initial CAPEX", "EUR/kW", 1.0),
            ("DIESEL_REPLACEMENT_COST", "Replacement", "EUR/kW", 1.0),
            ("DIESEL_VARIABLE_OM", "Variable O&M", "EUR/kWh_el", 1.0),
            ("DIESEL_OM_PER_OPERATING_HOUR", "Operating O&M", "EUR/h", 1.0),
            ("DIESEL_OM_PER_L", "Fuel-linked O&M", "EUR/L", 1.0),
            ("DIESEL_FUEL_COST", "Fuel cost", "EUR/L", 1.0),
        ]),
    ]

    all_specs = [spec for _, specs in groups for spec in specs]
    cost_vars = {key: tk.StringVar() for key, _, _, _ in all_specs}

    gui_keys = set(cost_vars)
    scenario_keys = set(COST_SCENARIOS[initial_scenario])
    if gui_keys != scenario_keys:
        missing = sorted(scenario_keys - gui_keys)
        extra = sorted(gui_keys - scenario_keys)
        root.destroy()
        raise RuntimeError(
            "Economic GUI and COST_SCENARIOS are inconsistent. "
            f"Missing GUI keys: {missing}; extra GUI keys: {extra}"
        )

    # Fixed action bar: always visible, independently of scroll position.
    button_frame = tk.Frame(root, padx=14, pady=8)
    button_frame.pack(side="bottom", fill="x", pady=(0, 4))

    # Scrollable content area.
    scroll_host = tk.Frame(root)
    scroll_host.pack(side="top", fill="both", expand=True)

    canvas = tk.Canvas(scroll_host, highlightthickness=0, borderwidth=0)
    vertical_scrollbar = tk.Scrollbar(
        scroll_host, orient="vertical", command=canvas.yview
    )
    canvas.configure(yscrollcommand=vertical_scrollbar.set)

    vertical_scrollbar.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)

    main = tk.Frame(canvas, padx=14, pady=12)
    main_window = canvas.create_window((0, 0), window=main, anchor="nw")

    def _update_scroll_region(_event=None):
        canvas.configure(scrollregion=canvas.bbox("all"))

    def _fit_content_width(event):
        # Make the inner frame follow the visible canvas width so no horizontal
        # scrolling is needed in normal use.
        canvas.itemconfigure(main_window, width=event.width)

    main.bind("<Configure>", _update_scroll_region)
    canvas.bind("<Configure>", _fit_content_width)

    def _on_mousewheel(event):
        # Windows/macOS mouse wheel.
        if getattr(event, "delta", 0):
            canvas.yview_scroll(int(-event.delta / 120), "units")

    def _on_linux_scroll_up(_event):
        canvas.yview_scroll(-1, "units")

    def _on_linux_scroll_down(_event):
        canvas.yview_scroll(1, "units")

    root.bind_all("<MouseWheel>", _on_mousewheel)
    root.bind_all("<Button-4>", _on_linux_scroll_up)
    root.bind_all("<Button-5>", _on_linux_scroll_down)

    tk.Label(
        main,
        text="Economic source scenario",
        font=("Arial", 11, "bold"),
    ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 5))

    scenario_frame = tk.Frame(main)
    scenario_frame.grid(row=1, column=0, columnspan=2, sticky="w")
    for idx, key in enumerate(COST_SCENARIOS):
        tk.Radiobutton(
            scenario_frame,
            text=COST_SCENARIO_LABELS.get(key, key),
            value=key,
            variable=scenario_var,
        ).grid(row=idx // 2, column=idx % 2, sticky="w", padx=(0, 18), pady=2)

    tk.Label(
        main,
        textvariable=description_var,
        justify="left",
        fg="gray35",
        wraplength=900,
    ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(5, 10))

    # Put the component frames in two columns.
    start_row = 3
    for idx, (group_name, specs) in enumerate(groups):
        col = idx % 2
        row = start_row + idx // 2
        frame = tk.LabelFrame(main, text=group_name, padx=10, pady=7)
        frame.grid(
            row=row, column=col, sticky="nsew",
            padx=(0, 8) if col == 0 else (8, 0), pady=5
        )

        tk.Label(frame, text="Parameter", font=("Arial", 9, "bold")).grid(
            row=0, column=0, sticky="w"
        )
        tk.Label(frame, text="Value", font=("Arial", 9, "bold")).grid(
            row=0, column=1, sticky="w"
        )
        tk.Label(frame, text="Unit", font=("Arial", 9, "bold")).grid(
            row=0, column=2, sticky="w"
        )

        for r, (key, label, unit, _scale) in enumerate(specs, start=1):
            tk.Label(frame, text=label + ":").grid(
                row=r, column=0, sticky="e", padx=(0, 8), pady=2
            )
            tk.Entry(frame, textvariable=cost_vars[key], width=13).grid(
                row=r, column=1, sticky="w", pady=2
            )
            tk.Label(frame, text=unit).grid(
                row=r, column=2, sticky="w", padx=(6, 0), pady=2
            )

    def load_cost_scenario(*_):
        scenario = scenario_var.get().upper()
        values = COST_SCENARIOS[scenario]
        description_var.set(COST_SCENARIO_DESCRIPTIONS.get(scenario, ""))
        for key, _label, _unit, scale in all_specs:
            cost_vars[key].set(f"{float(values[key]) * scale:g}")

    scenario_var.trace_add("write", load_cost_scenario)
    load_cost_scenario()

    # ``button_frame`` is intentionally outside the scrollable area so the
    # action buttons remain visible even when the cost table is taller than the
    # screen.

    def accept():
        try:
            manual_costs = {
                key: float(cost_vars[key].get()) / scale
                for key, _label, _unit, scale in all_specs
            }
        except ValueError:
            messagebox.showerror(
                "Invalid input",
                "All economic inputs must be numeric.",
                parent=root,
            )
            return

        if any(value < 0 for value in manual_costs.values()):
            messagebox.showerror(
                "Invalid input",
                "CAPEX, OPEX and replacement values cannot be negative.",
                parent=root,
            )
            return

        accepted["value"] = True
        accepted["costs"] = manual_costs
        accepted["scenario"] = scenario_var.get().upper()
        root.unbind_all("<MouseWheel>")
        root.unbind_all("<Button-4>")
        root.unbind_all("<Button-5>")
        root.destroy()

    def cancel():
        root.unbind_all("<MouseWheel>")
        root.unbind_all("<Button-4>")
        root.unbind_all("<Button-5>")
        root.destroy()

    tk.Button(
        button_frame, text="Continue", width=15, command=accept
    ).pack(side="right", padx=(6, 0))
    tk.Button(
        button_frame, text="Cancel", width=12, command=cancel
    ).pack(side="right")

    root.protocol("WM_DELETE_WINDOW", cancel)
    root.mainloop()

    if not accepted["value"]:
        raise SystemExit("Simulation cancelled by user.")

    chosen = accepted["scenario"]
    preset = COST_SCENARIOS[chosen]
    manual = accepted["costs"]
    overridden = any(
        abs(float(manual[k]) - float(preset[k])) > 1e-12
        for k in preset
    )
    source_label = COST_SCENARIO_LABELS.get(chosen, chosen)
    label = (
        source_label + " + MANUAL OVERRIDES"
        if overridden else source_label
    )
    apply_cost_inputs(
        manual,
        scenario_label=label,
        scenario_key=chosen,
    )

    el_specific = electrolyzer_specific_capex(EL_SIZE)
    fc_specific = fuel_cell_specific_capex(FC_SIZE)

    print("\nECONOMIC INPUTS USED FOR THIS RUN")
    print(f"  Scenario:                         {COST_SCENARIO}")
    print(f"  Financial horizon:                {PROJECT_LIFETIME:d} y, {100*DISCOUNT_RATE:.2f}% real discount rate")
    print(f"  PV CAPEX:                         {PV_CAPEX:.2f} EUR/kW")
    print(f"  PV O&M fixed / variable:          {PV_FIXED_OM:.4f} EUR/(kW y) | {PV_VARIABLE_OM:.6f} EUR/kWh")
    print(f"  Wind CAPEX:                       {WIND_CAPEX:.2f} EUR/kW")
    print(f"  Wind O&M fixed / variable:        {WIND_FIXED_OM:.4f} EUR/(kW y) | {WIND_VARIABLE_OM:.6f} EUR/kWh")
    print(f"  Battery CAPEX / replacement:      {BATTERY_CAPEX:.2f} | {BATTERY_REPLACEMENT_COST:.2f} EUR/kWh")
    print(f"  Battery O&M fixed / throughput:   {BATTERY_FIXED_OM:.4f} EUR/(kWh y) | {BATTERY_VARIABLE_OM:.6f} EUR/kWh")
    print(f"  Electrolyzer input CAPEX value:   {EL_CAPEX:.2f} EUR/kW")
    print(f"  Electrolyzer effective CAPEX:     {el_specific:.2f} EUR/kW at {EL_SIZE:.2f} kW")
    print(f"  EL stack replacement / frac O&M:  {100*EL_STACK_REPLACEMENT_FRACTION:.2f}% | {100*EL_TOTAL_OM_FRACTION:.2f}%/y")
    print(f"  EL absolute fixed O&M:             {EL_FIXED_OM_PER_KW_YEAR:.4f} EUR/(kW y)")
    print(f"  H2 tank CAPEX / replacement:      {H2_TANK_CAPEX:.2f} | {H2_TANK_REPLACEMENT_COST:.2f} EUR/kg")
    print(f"  H2 tank O&M:                       {100*H2_TANK_FIXED_OM_FRACTION:.2f}% CAPEX/y + {H2_TANK_FIXED_OM_PER_KG_YEAR:.4f} EUR/(kg y)")
    print(f"  Water treatment / compressor:     {H2_WATER_TREATMENT_CAPEX_PER_KW_EL:.2f} | {H2_COMPRESSOR_CAPEX_PER_KW_EL:.2f} EUR/kW_EL")
    print(f"  H2-system throughput O&M:         {H2_SYSTEM_VARIABLE_OM:.6f} EUR/kWh_H2 throughput")
    print(f"  Fuel-cell input CAPEX value:      {FC_CAPEX:.2f} EUR/kW")
    print(f"  Fuel-cell effective CAPEX:        {fc_specific:.2f} EUR/kW at {FC_SIZE:.2f} kW")
    print(f"  FC stack replacement / frac O&M:  {100*FC_STACK_REPLACEMENT_FRACTION:.2f}% | {100*FC_TOTAL_OM_FRACTION:.2f}%/y")
    print(f"  FC fixed / operating O&M:         {FC_FIXED_OM_PER_KW_YEAR:.4f} EUR/(kW y) | {FC_OM_PER_OPERATING_HOUR:.6f} EUR/op.h")
    print(f"  Diesel initial / replacement:     {DIESEL_CAPEX:.2f} | {DIESEL_REPLACEMENT_COST:.2f} EUR/kW")
    print(f"  Diesel O&M kWh / h / L:           {DIESEL_VARIABLE_OM:.6f} | {DIESEL_OM_PER_OPERATING_HOUR:.6f} | {DIESEL_OM_PER_L:.6f}")
    print(f"  Diesel fuel price:                {DIESEL_FUEL_COST:.2f} EUR/L")

def apply_system_mode(mode):
    """Attiva/disattiva i componenti in base alla configurazione scelta."""
    global PV_SIZE, N_WT, WIND_SIZE, FC_SIZE, EL_SIZE
    global H2_STORAGE_SIZE, BATTERY_SIZE, DIESEL_SIZE

    PV_SIZE = float(_BASELINE_SIZES["PV_SIZE"])
    N_WT = int(_BASELINE_SIZES["N_WT"])
    FC_SIZE = float(_BASELINE_SIZES["FC_SIZE"])
    EL_SIZE = float(_BASELINE_SIZES["EL_SIZE"])
    H2_STORAGE_SIZE = float(_BASELINE_SIZES["H2_STORAGE_SIZE"])
    BATTERY_SIZE = float(_BASELINE_SIZES["BATTERY_SIZE"])
    DIESEL_SIZE = float(_BASELINE_SIZES["DIESEL_SIZE"])
    WIND_SIZE = N_WT * WT_RATED_POWER

    if mode == "diesel_only":
        PV_SIZE = 0.0
        N_WT = 0
        WIND_SIZE = 0.0
        FC_SIZE = 0.0
        EL_SIZE = 0.0
        H2_STORAGE_SIZE = 0.0
        BATTERY_SIZE = 0.0
        DIESEL_SIZE = float(DIESEL_UNIT_SIZE)
        label = "DIESEL ONLY"

    elif mode == "renewable_only":
        DIESEL_SIZE = 0.0
        label = "100% RENEWABLE — BATTERY + H2 / NO DIESEL"

    elif mode == "mix":
        label = "HYBRID MIX — BATTERY + H2 + DIESEL"

    elif mode == "battery_only_mix":
        # Renewable generation + battery + existing diesel backup.
        # Hydrogen chain is completely disabled.
        FC_SIZE = 0.0
        EL_SIZE = 0.0
        H2_STORAGE_SIZE = 0.0
        label = "HYBRID MIX — BATTERY ONLY + DIESEL"

    elif mode == "battery_only_renewable":
        # Renewable generation + battery only. No hydrogen and no diesel.
        FC_SIZE = 0.0
        EL_SIZE = 0.0
        H2_STORAGE_SIZE = 0.0
        DIESEL_SIZE = 0.0
        label = "100% RENEWABLE — BATTERY ONLY / NO DIESEL"

    else:
        raise ValueError(f"Unknown SYSTEM_MODE: {mode}")

    print(f"\nSYSTEM CONFIGURATION: {label}")
    print(f"  PV:              {PV_SIZE:.2f} kW")
    print(f"  Wind turbines:   {N_WT:d} x {WT_RATED_POWER:.3f} kW")
    print(f"  Battery:         {BATTERY_SIZE:.2f} kWh")
    print(f"  Electrolyzer:    {EL_SIZE:.2f} kW")
    print(f"  H2 storage:      {H2_STORAGE_SIZE:.2f} kWh_H2")
    print(f"  Fuel cell:       {FC_SIZE:.2f} kW")
    print(f"  Diesel:          {DIESEL_SIZE:.2f} kW")


def select_excel_file():
    """Apre una finestra e restituisce il file Excel scelto dall'utente."""
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        selected_file = filedialog.askopenfilename(
            parent=root,
            title="Seleziona il file Excel con i dati di input",
            filetypes=[
                ("File Excel", "*.xlsx"),
                ("Tutti i file", "*.*"),
            ],
        )
    finally:
        root.destroy()

    if not selected_file:
        raise SystemExit("Nessun file selezionato. Simulazione annullata.")

    return Path(selected_file)


def select_plot_options():
    """Chiede quali grafici visualizzare dopo la simulazione.

    L'utente puo' scegliere indipendentemente:
      - vista annuale sintetica;
      - una settimana specifica (1-52);
      - un mese specifico (1-12).

    Restituisce un dizionario con le scelte, oppure None se si seleziona
    "Skip plots" / si chiude la finestra.
    """
    root = tk.Tk()
    root.title("Plot selection")
    root.resizable(False, False)
    root.attributes("-topmost", True)

    annual_var = tk.BooleanVar(value=PLOT_ANNUAL_DEFAULT)
    week_var = tk.BooleanVar(value=False)
    month_var = tk.BooleanVar(value=False)
    week_number_var = tk.StringVar(value=str(PLOT_WEEK_NUMBER))
    month_number_var = tk.StringVar(value=str(PLOT_MONTH_NUMBER))
    selection = {"value": None}

    frame = tk.Frame(root, padx=18, pady=16)
    frame.pack(fill="both", expand=True)

    tk.Label(
        frame,
        text="Choose which plots to display",
        font=("Arial", 12, "bold"),
    ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 12))

    tk.Checkbutton(
        frame,
        text="Whole year (full 8760-hour view)",
        variable=annual_var,
    ).grid(row=1, column=0, columnspan=3, sticky="w", pady=3)

    tk.Checkbutton(
        frame,
        text="Specific week",
        variable=week_var,
    ).grid(row=2, column=0, sticky="w", pady=3)
    tk.Label(frame, text="Week number:").grid(row=2, column=1, sticky="e", padx=(16, 6))
    week_spin = tk.Spinbox(
        frame,
        from_=1,
        to=52,
        width=6,
        textvariable=week_number_var,
    )
    week_spin.grid(row=2, column=2, sticky="w")

    tk.Checkbutton(
        frame,
        text="Specific month",
        variable=month_var,
    ).grid(row=3, column=0, sticky="w", pady=3)
    tk.Label(frame, text="Month number:").grid(row=3, column=1, sticky="e", padx=(16, 6))
    month_spin = tk.Spinbox(
        frame,
        from_=1,
        to=12,
        width=6,
        textvariable=month_number_var,
    )
    month_spin.grid(row=3, column=2, sticky="w")

    tk.Label(
        frame,
        text=(
            "By default only the full-year view is selected. "
            "You can also add a specific week and/or month. "
            "All views use the original hourly resolution."
        ),
        justify="left",
        fg="gray35",
    ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(10, 4))

    button_frame = tk.Frame(frame)
    button_frame.grid(row=5, column=0, columnspan=3, sticky="e", pady=(14, 0))

    def accept():
        try:
            week_number = int(week_number_var.get())
            month_number = int(month_number_var.get())
        except ValueError:
            messagebox.showerror(
                "Invalid plot selection",
                "Week and month numbers must be integers.",
                parent=root,
            )
            return

        if week_var.get() and not 1 <= week_number <= 52:
            messagebox.showerror(
                "Invalid week",
                "Week number must be between 1 and 52.",
                parent=root,
            )
            return
        if month_var.get() and not 1 <= month_number <= 12:
            messagebox.showerror(
                "Invalid month",
                "Month number must be between 1 and 12.",
                parent=root,
            )
            return

        selection["value"] = {
            "annual": bool(annual_var.get()),
            "week": week_number if week_var.get() else None,
            "month": month_number if month_var.get() else None,
        }
        root.destroy()

    def skip():
        selection["value"] = None
        root.destroy()

    tk.Button(button_frame, text="Show plots", width=12, command=accept).pack(
        side="right", padx=(6, 0)
    )
    tk.Button(button_frame, text="Skip plots", width=12, command=skip).pack(
        side="right"
    )

    root.protocol("WM_DELETE_WINDOW", skip)
    root.mainloop()
    return selection["value"]


def select_results_excel_file():
    """Apre la classica finestra 'Salva con nome' per i risultati Excel.

    Il file NON viene salvato automaticamente. L'utente sceglie cartella e nome.
    Se la finestra viene annullata, i risultati non vengono esportati.
    """
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    # Se disponibile, propone come cartella iniziale quella del file di input.
    initial_dir = None
    if EXCEL_FILE is not None:
        try:
            initial_dir = str(Path(EXCEL_FILE).resolve().parent)
        except Exception:
            initial_dir = None

    try:
        selected_file = filedialog.asksaveasfilename(
            parent=root,
            title="Salva i risultati della simulazione",
            initialdir=initial_dir,
            initialfile=RESULTS_EXCEL_FILE,
            defaultextension=".xlsx",
            filetypes=[
                ("File Excel", "*.xlsx"),
                ("Tutti i file", "*.*"),
            ],
        )
    finally:
        root.destroy()

    if not selected_file:
        return None

    return Path(selected_file)


def load_profiles(excel_file, sheet_name):
    """Legge B6:D8765. Il PV e il wind sono gia capacity factor."""
    global EXCEL_FILE
    if excel_file is None:
        excel_file = select_excel_file()
    # Memorizza esattamente il percorso e il nome del file scelto dall'utente.
    # Non viene imposto alcun nome predefinito al file Excel.
    EXCEL_FILE = Path(excel_file)
    raw = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, usecols="B:D",
                        skiprows=5, nrows=8760, header=None)
    raw.columns = ["load", "cf_wind", "cf_pv"]
    raw = raw.apply(pd.to_numeric, errors="coerce").dropna()
    if len(raw) != 8760:
        raise ValueError(f"Lette {len(raw)} righe valide, invece di 8760.")
    if (raw < 0).any().any():
        raise ValueError("Il file Excel contiene valori negativi.")
    return (raw["load"].to_numpy(), raw["cf_wind"].to_numpy(),
            raw["cf_pv"].to_numpy())


def simulate(load, cf_wind, cf_pv):
    """Priorita di dispatch: RES, batteria, fuel cell, diesel, non-servito."""
    n = len(load)
    diesel_size = (
        float(np.max(load)) if DIESEL_SIZE is None else float(DIESEL_SIZE)
    )
    if diesel_size < 0:
        raise ValueError("DIESEL_SIZE cannot be negative.")
    # Limiti di funzionamento delle macchine
    fc_min, fc_max = FC_MIN_LOAD_FRACTION * FC_SIZE, FC_SIZE
    el_min, el_max = EL_MIN_LOAD_FRACTION * EL_SIZE, EL_SIZE
    dg_min = DIESEL_MIN_LOAD_FRACTION * diesel_size
    dg_max = diesel_size
    h2_soc_min = H2_MIN_PRESSURE / H2_MAX_PRESSURE

    pv_power = cf_pv * PV_SIZE
    wind_power = cf_wind * WIND_SIZE
    res_power = pv_power + wind_power
    names = ["res_to_load", "battery_to_load", "fc_to_load", "diesel_to_load",
             "unserved", "res_to_battery", "dispatch_to_battery", "res_to_el",
             "curtailment", "dispatch_curtailment",
             "battery_power", "fc_power", "el_power",
             "el_efficiency", "fc_efficiency",
             "h2_produced", "h2_consumed",
             "diesel_power", "diesel_fuel",
             "diesel_shaft_power", "alternator_loss"]
    out = {name: np.zeros(n) for name in names}
    bat = np.zeros(n + 1)
    h2 = np.zeros(n + 1)
    bat[0] = BATTERY_SIZE * BATTERY_SOC_INITIAL
    h2[0] = H2_STORAGE_SIZE * H2_SOC_INITIAL

    # Potenza nominale all'albero del motore: la curva di consumo e' riferita
    # ad essa, mentre diesel_size e' la potenza elettrica al quadro.
    diesel_shaft_rated = (
        diesel_size / alternator_efficiency(diesel_size)
        if diesel_size > 0 else 0.0
    )

    # Curva di consumo diesel dell'originale MATLAB (litri/ora)
    start_fuel = (DIESEL_START_DURATION / 60) * (
        DIESEL_FUEL_CURVE_A + DIESEL_FUEL_CURVE_B
    ) * diesel_shaft_rated

    for t in range(n):
        bat_old = bat[t] * (1 - BATTERY_SELF_DISCHARGE)
        h2_old = h2[t]
        bat_min = BATTERY_SIZE * BATTERY_SOC_MIN
        bat_max = BATTERY_SIZE * BATTERY_SOC_MAX
        h2_min = H2_STORAGE_SIZE * h2_soc_min
        h2_max = H2_STORAGE_SIZE

        if res_power[t] >= load[t]:
            # Surplus: prima batteria, poi elettrolizzatore, poi curtailment.
            out["res_to_load"][t] = load[t]
            surplus = res_power[t] - load[t]
            max_charge_input = max(0, (bat_max - bat_old) / BATTERY_CHARGE_EFFICIENCY)
            out["res_to_battery"][t] = min(surplus, max_charge_input)
            # Battery power at the AC busbar: positive means battery charging.
            out["battery_power"][t] = out["res_to_battery"][t]
            # Battery energy changes after the charging-converter and battery losses.
            bat_new = bat_old + out["res_to_battery"][t] * BATTERY_CHARGE_EFFICIENCY
            surplus -= out["res_to_battery"][t]
            h2_new = h2_old
            if surplus >= el_min and EL_SIZE > 0:
                # Con efficienza variabile, il limite imposto dallo spazio
                # residuo nel serbatoio non e' semplicemente room / eta_const.
                max_el_from_h2_room = max_electrolyzer_power_for_h2_room(
                    h2_max - h2_old, EL_SIZE
                )
                out["res_to_el"][t] = min(
                    surplus, el_max, max_el_from_h2_room
                )
                if out["res_to_el"][t] + 1e-12 < el_min:
                    out["res_to_el"][t] = 0.0

                out["el_power"][t] = out["res_to_el"][t]
                if out["el_power"][t] > 0:
                    eta_el = electrolyzer_efficiency(out["el_power"][t], EL_SIZE)
                    h2_produced = out["el_power"][t] * eta_el  # kWh_H2 in 1 h
                    out["el_efficiency"][t] = eta_el
                    out["h2_produced"][t] = h2_produced
                    h2_new += h2_produced
                    surplus -= out["res_to_el"][t]
            out["curtailment"][t] = surplus
        else:
            # Deficit: prima batteria, poi fuel cell, poi diesel.
            out["res_to_load"][t] = res_power[t]
            total_deficit = load[t] - res_power[t]
            eps = 1e-9

            # Prima assegnazione secondo la priorita: batteria, FC, diesel.
            battery_to_load = min(
                total_deficit,
                max(0, (bat_old - bat_min) * BATTERY_DISCHARGE_EFFICIENCY)
            )
            remaining = total_deficit - battery_to_load

            # Maximum FC output is obtained by solving
            # P_FC / eta_FC(P_FC) <= available H2 energy.
            fc_available = min(
                fc_max,
                max_fuel_cell_power_from_h2(h2_old - h2_min, FC_SIZE)
            )
            fc_output = 0.0
            if remaining > eps and fc_available >= fc_min:
                # Se la richiesta e inferiore al minimo tecnico, la FC lavora
                # comunque al minimo; l'eccesso verra redistribuito sotto.
                fc_output = min(max(remaining, fc_min), fc_available)

            remaining = total_deficit - battery_to_load - fc_output
            diesel_output = 0.0
            if remaining > eps:
                diesel_output = min(dg_max, max(dg_min, remaining))

            # Un minimo tecnico puo creare sovrapproduzione. Quando e il diesel
            # a causarla, esso sostituisce prima la FC e poi la batteria. La FC
            # puo essere spenta, oppure deve restare almeno a fc_min.
            excess = max(0.0, battery_to_load + fc_output + diesel_output - total_deficit)
            if diesel_output > 0 and excess > eps and fc_output > 0:
                if excess + eps >= fc_output:
                    excess -= fc_output
                    fc_output = 0.0
                elif fc_output - excess + eps >= fc_min:
                    fc_output -= excess
                    excess = 0.0
                else:
                    reducible_fc = max(0.0, fc_output - fc_min)
                    fc_output = fc_min
                    excess -= reducible_fc

            # L'eccesso ancora presente annulla progressivamente la scarica.
            if excess > eps and battery_to_load > 0:
                battery_reduction = min(excess, battery_to_load)
                battery_to_load -= battery_reduction
                excess -= battery_reduction

            # Solo dopo avere annullato la scarica, l'eventuale energia residua
            # puo caricare la batteria; il resto e produzione non utilizzabile.
            dispatch_to_battery = 0.0
            if excess > eps:
                max_charge_input = max(
                    0, (bat_max - bat_old) / BATTERY_CHARGE_EFFICIENCY
                )
                dispatch_to_battery = min(excess, max_charge_input)
                excess -= dispatch_to_battery

            out["battery_to_load"][t] = battery_to_load
            out["dispatch_to_battery"][t] = dispatch_to_battery
            out["battery_power"][t] = dispatch_to_battery - battery_to_load
            out["fc_power"][t] = fc_output
            out["diesel_power"][t] = diesel_output

            # Le potenze verso il carico escludono la quota destinata alla
            # batteria o dissipata per rispettare i minimi tecnici.
            load_after_battery = max(0.0, total_deficit - battery_to_load)
            out["fc_to_load"][t] = min(fc_output, load_after_battery)
            load_after_fc = max(0.0, load_after_battery - out["fc_to_load"][t])
            out["diesel_to_load"][t] = min(diesel_output, load_after_fc)
            out["dispatch_curtailment"][t] = max(0.0, excess)

            served_load = (battery_to_load + out["fc_to_load"][t] + out["diesel_to_load"][t])
            out["unserved"][t] = max(0.0, total_deficit - served_load)

            # Gli stati vengono aggiornati una sola volta, usando il dispatch
            # definitivo dopo tutti i check sui minimi tecnici.
            bat_new = (
                bat_old
                - battery_to_load / BATTERY_DISCHARGE_EFFICIENCY
                + dispatch_to_battery * BATTERY_CHARGE_EFFICIENCY
            )
            h2_new = h2_old
            if fc_output > 0:
                eta_fc = fuel_cell_efficiency(fc_output, FC_SIZE)
                h2_consumed = fc_output / eta_fc  # kWh_H2 in 1 h
                out["fc_efficiency"][t] = eta_fc
                out["h2_consumed"][t] = h2_consumed
                h2_new -= h2_consumed

            if diesel_output > 0:
                # Potenza all'albero = potenza elettrica / rendimento alternatore.
                # La curva di consumo e' funzione della potenza all'albero.
                diesel_shaft = diesel_output / alternator_efficiency(diesel_output)
                out["diesel_shaft_power"][t] = diesel_shaft
                out["alternator_loss"][t] = diesel_shaft - diesel_output
                out["diesel_fuel"][t] = (
                    DIESEL_FUEL_CURVE_B * diesel_shaft_rated
                    + DIESEL_FUEL_CURVE_A * diesel_shaft
                )
                if t == 0 or out["diesel_power"][t-1] == 0:
                    out["diesel_fuel"][t] += start_fuel

        bat[t+1] = np.clip(bat_new, bat_min, bat_max)
        h2[t+1] = np.clip(h2_new, h2_min, h2_max)

    results = pd.DataFrame(out)
    results.insert(0, "load", load)
    results.insert(1, "pv", pv_power)
    results.insert(2, "wind", wind_power)
    results.insert(3, "res_total", res_power)
    # Divisione sicura: quando batteria o serbatoio H2 sono disattivati
    # la capacita' e' zero, quindi il SOC resta a zero senza divisioni per zero.
    results["battery_soc"] = (
        bat[:-1] / BATTERY_SIZE if BATTERY_SIZE > 0 else np.zeros(n)
    )
    results["h2_soc"] = (
        h2[:-1] / H2_STORAGE_SIZE if H2_STORAGE_SIZE > 0 else np.zeros(n)
    )
    # Save the true boundary states as attributes. ``battery_soc`` above contains
    # the state at the beginning of each simulated hour (bat[:-1]); bat[-1] is the
    # actual state after hour 8760 and is therefore the correct terminal SOC.
    results.attrs["battery_soc_initial"] = (
        float(BATTERY_SOC_INITIAL) if BATTERY_SIZE > 0 else 0.0
    )
    results.attrs["battery_soc_final"] = (
        float(bat[-1] / BATTERY_SIZE) if BATTERY_SIZE > 0 else 0.0
    )
    results.attrs["h2_soc_initial"] = (
        float(H2_SOC_INITIAL) if H2_STORAGE_SIZE > 0 else 0.0
    )
    results.attrs["h2_soc_final"] = (
        float(h2[-1] / H2_STORAGE_SIZE) if H2_STORAGE_SIZE > 0 else 0.0
    )
    # Explicit LOH aliases for the annual hydrogen boundary condition.
    results.attrs["h2_loh_initial"] = results.attrs["h2_soc_initial"]
    results.attrs["h2_loh_final"] = results.attrs["h2_soc_final"]
    results.attrs["battery_terminal_soc_feasible"] = (
        True if BATTERY_SIZE <= 0 else
        results.attrs["battery_soc_final"] >=
        results.attrs["battery_soc_initial"] - BATTERY_SOC_END_TOLERANCE
    )
    results.attrs["h2_terminal_loh_feasible"] = (
        True if H2_STORAGE_SIZE <= 0 else
        results.attrs["h2_loh_final"] >=
        results.attrs["h2_loh_initial"] - H2_LOH_END_TOLERANCE
    )
    results.attrs["storage_terminal_feasible"] = (
        results.attrs["battery_terminal_soc_feasible"]
        and results.attrs["h2_terminal_loh_feasible"]
    )
    results.attrs["diesel_size"] = diesel_size
    results.attrs["diesel_shaft_rated"] = diesel_shaft_rated
    return results


def count_starts(power, tolerance=1e-9):
    """Conta gli avviamenti: passaggi da macchina spenta a macchina accesa."""
    on = np.asarray(power) > tolerance
    if len(on) == 0:
        return 0
    return int(on[0]) + int(np.count_nonzero(on[1:] & ~on[:-1]))


def replacement_cash_flow(replacement_cost, lifetime, discount_rate,
                          project_lifetime, include_salvage):
    """Valore attuale di sostituzioni e valore residuo secondo il MATLAB."""
    if replacement_cost <= 0 or lifetime <= 0:
        return 0.0, 0.0, []

    replacement_present_value = 0.0
    replacement_years = []
    next_replacement_age = lifetime

    # Il MATLAB non effettua una sostituzione esattamente all'ultimo anno.
    while math.ceil(next_replacement_age) < project_lifetime:
        replacement_year = math.ceil(next_replacement_age)
        replacement_years.append(replacement_year)
        replacement_present_value += (
            replacement_cost / (1 + discount_rate) ** replacement_year
        )
        next_replacement_age += lifetime

    salvage_present_value = 0.0
    if (include_salvage and replacement_years
            and next_replacement_age > project_lifetime):
        last_replacement_age = next_replacement_age - lifetime
        component_age_at_project_end = project_lifetime - last_replacement_age
        remaining_fraction = np.clip(
            (lifetime - component_age_at_project_end) / lifetime,
            0.0, 1.0
        )
        salvage_present_value = (
            replacement_cost * remaining_fraction
            / (1 + discount_rate) ** project_lifetime
        )

    return (
        replacement_present_value,
        float(salvage_present_value),
        replacement_years,
    )


def calculate_operational_metrics(r):
    """Calcola utilizzo, consumo di vita e vita residua dei componenti."""
    tolerance = 1e-9
    fc_operating_time = int(
        np.count_nonzero(r["fc_power"].to_numpy() > tolerance)
    )  # h/year
    fc_starts = count_starts(r["fc_power"], tolerance)
    el_operating_time = int(
        np.count_nonzero(r["el_power"].to_numpy() > tolerance)
    )  # h/year
    el_starts = count_starts(r["el_power"], tolerance)
    diesel_operating_time = int(
        np.count_nonzero(r["diesel_power"].to_numpy() > tolerance)
    )  # h/year

    # Hydrogen-system conversion KPIs. These are energy-weighted yearly
    # efficiencies, consistent with the hourly part-load curves.
    el_input_energy = float(r["el_power"].sum())  # kWh_el/year
    h2_produced_energy = float(r["h2_produced"].sum())  # kWh_H2,LHV/year
    h2_consumed_energy = float(r["h2_consumed"].sum())  # kWh_H2,LHV/year
    fc_output_energy = float(r["fc_power"].sum())  # kWh_el/year

    el_average_efficiency = (
        h2_produced_energy / el_input_energy
        if el_input_energy > tolerance else np.nan
    )
    fc_average_efficiency = (
        fc_output_energy / h2_consumed_energy
        if h2_consumed_energy > tolerance else np.nan
    )
    h2_conversion_chain_efficiency = (
        el_average_efficiency * fc_average_efficiency
        if np.isfinite(el_average_efficiency) and np.isfinite(fc_average_efficiency)
        else np.nan
    )
    el_average_load_fraction = (
        el_input_energy / (EL_SIZE * el_operating_time)
        if EL_SIZE > 0 and el_operating_time > 0 else np.nan
    )
    fc_average_load_fraction = (
        fc_output_energy / (FC_SIZE * fc_operating_time)
        if FC_SIZE > 0 and fc_operating_time > 0 else np.nan
    )

    battery_operating_time = int(
        np.count_nonzero(
            np.abs(r["battery_power"].to_numpy()) > tolerance
        )
    )  # h/year

    battery_charge_internal = (
        r["res_to_battery"].sum() + r["dispatch_to_battery"].sum()
    ) * BATTERY_CHARGE_EFFICIENCY
    battery_discharge_internal = (
        r["battery_to_load"].sum() / BATTERY_DISCHARGE_EFFICIENCY
    )
    battery_throughput = battery_charge_internal + battery_discharge_internal

    battery_lifetime_throughput = (
        2 * BATTERY_SIZE * BATTERY_CTF_DOD_AVERAGE
    )

    battery_damage = (
        battery_throughput / battery_lifetime_throughput
        if battery_lifetime_throughput > 0 else 0.0
    )
    el_damage = (
        el_operating_time / EL_STACK_LIFETIME
        + el_starts / EL_STACK_START_LIFETIME
    )
    fc_damage = (
        fc_operating_time / FC_STACK_LIFETIME
        + fc_starts / FC_STACK_START_LIFETIME
    )
    diesel_damage = diesel_operating_time / DIESEL_LIFETIME

    battery_life = (
        1 / battery_damage if battery_damage > tolerance else PROJECT_LIFETIME
    )
    el_life = (
        1 / el_damage if el_damage > tolerance else PROJECT_LIFETIME
    )
    fc_life = (
        1 / fc_damage if fc_damage > tolerance else PROJECT_LIFETIME
    )
    diesel_life = (
        1 / diesel_damage if diesel_damage > tolerance else PROJECT_LIFETIME
    )

    return {
        "fc_operating_time": fc_operating_time,
        "fc_starts": fc_starts,
        "el_operating_time": el_operating_time,
        "el_starts": el_starts,
        "diesel_operating_time": diesel_operating_time,
        "battery_operating_time": battery_operating_time,
        "el_input_energy": el_input_energy,
        "h2_produced_energy": h2_produced_energy,
        "h2_consumed_energy": h2_consumed_energy,
        "fc_output_energy": fc_output_energy,
        "el_average_efficiency": el_average_efficiency,
        "fc_average_efficiency": fc_average_efficiency,
        "h2_conversion_chain_efficiency": h2_conversion_chain_efficiency,
        "el_average_load_fraction": el_average_load_fraction,
        "fc_average_load_fraction": fc_average_load_fraction,
        "battery_throughput": battery_throughput,
        "battery_lifetime_throughput": battery_lifetime_throughput,
        "battery_damage": battery_damage,
        "el_damage": el_damage,
        "fc_damage": fc_damage,
        "diesel_damage": diesel_damage,
        "battery_life": battery_life,
        "el_life": el_life,
        "fc_life": fc_life,
        "diesel_life": diesel_life,
        "battery_remaining_life": max(
            0.0, BATTERY_REMAINING_LIFE_INITIAL - battery_damage
        ),
        "el_remaining_life": max(
            0.0, EL_REMAINING_LIFE_INITIAL - el_damage
        ),
        "fc_remaining_life": max(
            0.0, FC_REMAINING_LIFE_INITIAL - fc_damage
        ),
        "diesel_remaining_life": max(
            0.0, DIESEL_REMAINING_LIFE_INITIAL - diesel_damage
        ),
    }


def calculate_economics(r, metrics=None):
    """Calcola KPI, costi attualizzati, NPC e LCOE."""
    diesel_size = r.attrs["diesel_size"]  # kW
    metrics = metrics or calculate_operational_metrics(r)
    discount_factor = sum(
        (1 + DISCOUNT_RATE) ** -year
        for year in range(1, PROJECT_LIFETIME + 1)
    )

    fc_operating_time = metrics["fc_operating_time"]
    fc_starts = metrics["fc_starts"]
    el_operating_time = metrics["el_operating_time"]
    el_starts = metrics["el_starts"]
    diesel_operating_time = metrics["diesel_operating_time"]
    battery_operating_time = metrics["battery_operating_time"]
    battery_throughput = metrics["battery_throughput"]
    battery_life = metrics["battery_life"]
    el_life = metrics["el_life"]
    fc_life = metrics["fc_life"]
    diesel_life = metrics["diesel_life"]

    el_specific_capex = electrolyzer_specific_capex(EL_SIZE)
    fc_specific_capex = fuel_cell_specific_capex(FC_SIZE)
    el_capex = el_specific_capex * EL_SIZE
    fc_capex = fc_specific_capex * FC_SIZE
    h2_mass = H2_STORAGE_SIZE * 3.6 / H2_LHV
    h2_aux_capex = h2_auxiliary_capex()
    h2_throughput = (
        metrics["h2_produced_energy"] + metrics["h2_consumed_energy"]
    )

    annual_costs = {
        "Wind": (
            WIND_FIXED_OM * WIND_SIZE
            + WIND_VARIABLE_OM * r["wind"].sum()
        ),
        "PV": (
            PV_FIXED_OM * PV_SIZE
            + PV_VARIABLE_OM * r["pv"].sum()
        ),
        "Battery Li-ion": (
            BATTERY_FIXED_OM * BATTERY_SIZE
            + BATTERY_VARIABLE_OM * battery_throughput
        ),
        "PEM electrolyzer": (
            EL_TOTAL_OM_FRACTION / 3 * el_capex
            + (2 / 3) * EL_TOTAL_OM_FRACTION
            * el_capex / 8760 * el_operating_time
            + EL_FIXED_OM_PER_KW_YEAR * EL_SIZE
        ),
        "PEM fuel cell": (
            FC_TOTAL_OM_FRACTION / 3 * fc_capex
            + (2 / 3) * FC_TOTAL_OM_FRACTION
            * fc_capex / 8760 * fc_operating_time
            + FC_FIXED_OM_PER_KW_YEAR * FC_SIZE
            + FC_OM_PER_OPERATING_HOUR * fc_operating_time
        ),
        "H2 tank": (
            H2_TANK_FIXED_OM_FRACTION * H2_TANK_CAPEX * h2_mass
            + H2_TANK_FIXED_OM_PER_KG_YEAR * h2_mass
        ),
        "H2 auxiliaries": H2_SYSTEM_VARIABLE_OM * h2_throughput,
        "Diesel generator": (
            DIESEL_FIXED_OM_FRACTION * DIESEL_CAPEX * diesel_size
            + DIESEL_VARIABLE_OM * r["diesel_power"].sum()
            + DIESEL_OM_PER_OPERATING_HOUR * diesel_operating_time
            + DIESEL_OM_PER_L * r["diesel_fuel"].sum()
            + DIESEL_FUEL_COST * r["diesel_fuel"].sum()
        ),
    }

    components = [
        ("Wind", WIND_CAPEX * WIND_SIZE,
         WIND_REPLACEMENT_FRACTION * WIND_CAPEX * WIND_SIZE,
         WIND_LIFETIME, False),
        ("PV", PV_CAPEX * PV_SIZE,
         PV_REPLACEMENT_FRACTION * PV_CAPEX * PV_SIZE,
         PV_LIFETIME, False),
        ("Battery Li-ion", BATTERY_CAPEX * BATTERY_SIZE,
         BATTERY_REPLACEMENT_COST * BATTERY_SIZE,
         battery_life, True),
        ("PEM electrolyzer", el_capex,
         EL_STACK_REPLACEMENT_FRACTION * el_capex,
         el_life, True),
        ("PEM fuel cell", fc_capex,
         FC_STACK_REPLACEMENT_FRACTION * fc_capex,
         fc_life, True),
        ("H2 tank", H2_TANK_CAPEX * h2_mass,
         H2_TANK_REPLACEMENT_COST * h2_mass,
         H2_TANK_LIFETIME, False),
        ("H2 auxiliaries", h2_aux_capex,
         0.0, PROJECT_LIFETIME, False),
        ("Diesel generator", DIESEL_CAPEX * diesel_size,
         DIESEL_REPLACEMENT_COST * diesel_size,
         diesel_life, True),
    ]

    cost_rows = []
    for component, capex, replacement_cost, lifetime, include_salvage in components:
        replacement_present_value, salvage_present_value, replacement_years = (
            replacement_cash_flow(
                replacement_cost, lifetime, DISCOUNT_RATE,
                PROJECT_LIFETIME, include_salvage
            )
        )
        opex_present_value = annual_costs[component] * discount_factor
        npc = (
            capex + opex_present_value
            + replacement_present_value - salvage_present_value
        )
        cost_rows.append({
            "Component": component,
            "Calculated lifetime [years]": lifetime,
            "Replacement years": ", ".join(map(str, replacement_years)) or "None",
            "CAPEX [M€]": capex / 1e6,
            "Annual OPEX [M€/year]": annual_costs[component] / 1e6,
            "OPEX present value [M€]": opex_present_value / 1e6,
            "Replacement present value [M€]": replacement_present_value / 1e6,
            "Salvage present value [M€]": salvage_present_value / 1e6,
            "NPC [M€]": npc / 1e6,
        })

    breakdown = pd.DataFrame(cost_rows)
    total_npc = breakdown["NPC [M€]"].sum()
    annual_served_energy = (r["load"] - r["unserved"]).sum()
    discounted_served_energy = annual_served_energy * discount_factor
    lcoe = (
        total_npc * 1e6 / discounted_served_energy
        if discounted_served_energy > 0 else np.nan
    )
    breakdown["LCOE contribution [€/kWh]"] = (
        breakdown["NPC [M€]"] * 1e6 / discounted_served_energy
        if discounted_served_energy > 0 else np.nan
    )

    kpis = {
        "Electrolyzer operating hours [h/year]": el_operating_time,
        "Electrolyzer starts [starts/year]": el_starts,
        "Electrolyzer stack lifetime [years]": el_life,
        "Electrolyzer average load fraction [-]": metrics["el_average_load_fraction"],
        "Electrolyzer average system efficiency [-]": metrics["el_average_efficiency"],
        "H2 produced [kWh_H2 LHV/year]": metrics["h2_produced_energy"],
        "Fuel cell operating hours [h/year]": fc_operating_time,
        "Fuel cell starts [starts/year]": fc_starts,
        "Fuel cell stack lifetime [years]": fc_life,
        "Fuel cell average load fraction [-]": metrics["fc_average_load_fraction"],
        "Fuel cell average system efficiency [-]": metrics["fc_average_efficiency"],
        "H2 consumed [kWh_H2 LHV/year]": metrics["h2_consumed_energy"],
        "H2 conversion-chain efficiency [-]": metrics["h2_conversion_chain_efficiency"],
        "Battery operating hours [h/year]": battery_operating_time,
        "Battery annual throughput [kWh/year]": battery_throughput,
        "Battery lifetime [years]": battery_life,
        "Diesel rated power [kW]": diesel_size,
        "Diesel operating hours [h/year]": diesel_operating_time,
        "Diesel lifetime [years]": diesel_life,
    }
    summary = {
        "Discount rate [%]": DISCOUNT_RATE * 100,
        "Present-worth factor [-]": discount_factor,
        "Annual served energy [MWh/year]": annual_served_energy / 1000,
        "Discounted lifetime served energy [MWh]": discounted_served_energy / 1000,
        "Total CAPEX [M€]": breakdown["CAPEX [M€]"].sum(),
        "Total OPEX present value [M€]": breakdown["OPEX present value [M€]"].sum(),
        "Total replacement present value [M€]": breakdown["Replacement present value [M€]"].sum(),
        "Total salvage present value [M€]": breakdown["Salvage present value [M€]"].sum(),
        "Total NPC [M€]": total_npc,
        "LCOE [€/kWh]": lcoe,
    }
    return {
        "summary": summary,
        "kpis": kpis,
        "breakdown": breakdown,
    }


def calculate_annual_operation(r, metrics=None):
    """Annual cash O&M plus degradation-reserve indicators."""
    metrics = metrics or calculate_operational_metrics(r)
    diesel_size = r.attrs["diesel_size"]

    el_capex = electrolyzer_specific_capex(EL_SIZE) * EL_SIZE
    fc_capex = fuel_cell_specific_capex(FC_SIZE) * FC_SIZE
    h2_mass = H2_STORAGE_SIZE * 3.6 / H2_LHV if H2_STORAGE_SIZE > 0 else 0.0
    h2_throughput = metrics["h2_produced_energy"] + metrics["h2_consumed_energy"]

    battery_replacement = BATTERY_REPLACEMENT_COST * BATTERY_SIZE
    el_stack_replacement = EL_STACK_REPLACEMENT_FRACTION * el_capex
    fc_stack_replacement = FC_STACK_REPLACEMENT_FRACTION * fc_capex
    diesel_replacement = DIESEL_REPLACEMENT_COST * diesel_size

    el_fixed_om = (
        EL_TOTAL_OM_FRACTION / 3 * el_capex
        + EL_FIXED_OM_PER_KW_YEAR * EL_SIZE
    )
    el_variable_om = (
        (2 / 3) * EL_TOTAL_OM_FRACTION * el_capex / 8760
        * metrics["el_operating_time"]
    )
    fc_fixed_om = (
        FC_TOTAL_OM_FRACTION / 3 * fc_capex
        + FC_FIXED_OM_PER_KW_YEAR * FC_SIZE
    )
    fc_variable_om = (
        (2 / 3) * FC_TOTAL_OM_FRACTION * fc_capex / 8760
        * metrics["fc_operating_time"]
        + FC_OM_PER_OPERATING_HOUR * metrics["fc_operating_time"]
    )

    cost_rows = [
        {
            "Component": "PV",
            "Fixed O&M [€/year]": PV_FIXED_OM * PV_SIZE,
            "Variable O&M [€/year]": PV_VARIABLE_OM * r["pv"].sum(),
            "Fuel [€/year]": 0.0,
            "Degradation from operation [€/year]": 0.0,
            "Degradation from starts [€/year]": 0.0,
        },
        {
            "Component": "Wind",
            "Fixed O&M [€/year]": WIND_FIXED_OM * WIND_SIZE,
            "Variable O&M [€/year]": WIND_VARIABLE_OM * r["wind"].sum(),
            "Fuel [€/year]": 0.0,
            "Degradation from operation [€/year]": 0.0,
            "Degradation from starts [€/year]": 0.0,
        },
        {
            "Component": "Battery Li-ion",
            "Fixed O&M [€/year]": BATTERY_FIXED_OM * BATTERY_SIZE,
            "Variable O&M [€/year]": (
                BATTERY_VARIABLE_OM * metrics["battery_throughput"]
            ),
            "Fuel [€/year]": 0.0,
            "Degradation from operation [€/year]": (
                battery_replacement * metrics["battery_damage"]
            ),
            "Degradation from starts [€/year]": 0.0,
        },
        {
            "Component": "PEM electrolyzer",
            "Fixed O&M [€/year]": el_fixed_om,
            "Variable O&M [€/year]": el_variable_om,
            "Fuel [€/year]": 0.0,
            "Degradation from operation [€/year]": (
                el_stack_replacement * metrics["el_operating_time"]
                / EL_STACK_LIFETIME
            ),
            "Degradation from starts [€/year]": (
                el_stack_replacement * metrics["el_starts"]
                / EL_STACK_START_LIFETIME
            ),
        },
        {
            "Component": "H2 tank",
            "Fixed O&M [€/year]": (
                H2_TANK_FIXED_OM_FRACTION * H2_TANK_CAPEX * h2_mass
                + H2_TANK_FIXED_OM_PER_KG_YEAR * h2_mass
            ),
            "Variable O&M [€/year]": 0.0,
            "Fuel [€/year]": 0.0,
            "Degradation from operation [€/year]": 0.0,
            "Degradation from starts [€/year]": 0.0,
        },
        {
            "Component": "H2 auxiliaries",
            "Fixed O&M [€/year]": 0.0,
            "Variable O&M [€/year]": H2_SYSTEM_VARIABLE_OM * h2_throughput,
            "Fuel [€/year]": 0.0,
            "Degradation from operation [€/year]": 0.0,
            "Degradation from starts [€/year]": 0.0,
        },
        {
            "Component": "PEM fuel cell",
            "Fixed O&M [€/year]": fc_fixed_om,
            "Variable O&M [€/year]": fc_variable_om,
            "Fuel [€/year]": 0.0,
            "Degradation from operation [€/year]": (
                fc_stack_replacement * metrics["fc_operating_time"]
                / FC_STACK_LIFETIME
            ),
            "Degradation from starts [€/year]": (
                fc_stack_replacement * metrics["fc_starts"]
                / FC_STACK_START_LIFETIME
            ),
        },
        {
            "Component": "Diesel generator",
            "Fixed O&M [€/year]": (
                DIESEL_FIXED_OM_FRACTION * DIESEL_CAPEX * diesel_size
            ),
            "Variable O&M [€/year]": (
                DIESEL_VARIABLE_OM * r["diesel_power"].sum()
                + DIESEL_OM_PER_OPERATING_HOUR * metrics["diesel_operating_time"]
                + DIESEL_OM_PER_L * r["diesel_fuel"].sum()
            ),
            "Fuel [€/year]": DIESEL_FUEL_COST * r["diesel_fuel"].sum(),
            "Degradation from operation [€/year]": (
                diesel_replacement * metrics["diesel_damage"]
            ),
            "Degradation from starts [€/year]": 0.0,
        },
    ]

    cost_breakdown = pd.DataFrame(cost_rows)
    cost_columns = [
        "Fixed O&M [€/year]",
        "Variable O&M [€/year]",
        "Fuel [€/year]",
        "Degradation from operation [€/year]",
        "Degradation from starts [€/year]",
    ]
    cost_breakdown["Total annual operating cost [€/year]"] = (
        cost_breakdown[cost_columns].sum(axis=1)
    )

    remaining_life = pd.DataFrame([
        {
            "Component": "Battery Li-ion",
            "Initial remaining life [%]": BATTERY_REMAINING_LIFE_INITIAL * 100,
            "Life consumption from operation [%/year]": metrics["battery_damage"] * 100,
            "Life consumption from starts [%/year]": 0.0,
            "End-of-year remaining life [%]": metrics["battery_remaining_life"] * 100,
        },
        {
            "Component": "PEM electrolyzer",
            "Initial remaining life [%]": EL_REMAINING_LIFE_INITIAL * 100,
            "Life consumption from operation [%/year]": (
                metrics["el_operating_time"] / EL_STACK_LIFETIME * 100
            ),
            "Life consumption from starts [%/year]": (
                metrics["el_starts"] / EL_STACK_START_LIFETIME * 100
            ),
            "End-of-year remaining life [%]": metrics["el_remaining_life"] * 100,
        },
        {
            "Component": "PEM fuel cell",
            "Initial remaining life [%]": FC_REMAINING_LIFE_INITIAL * 100,
            "Life consumption from operation [%/year]": (
                metrics["fc_operating_time"] / FC_STACK_LIFETIME * 100
            ),
            "Life consumption from starts [%/year]": (
                metrics["fc_starts"] / FC_STACK_START_LIFETIME * 100
            ),
            "End-of-year remaining life [%]": metrics["fc_remaining_life"] * 100,
        },
        {
            "Component": "Diesel generator",
            "Initial remaining life [%]": DIESEL_REMAINING_LIFE_INITIAL * 100,
            "Life consumption from operation [%/year]": metrics["diesel_damage"] * 100,
            "Life consumption from starts [%/year]": 0.0,
            "End-of-year remaining life [%]": metrics["diesel_remaining_life"] * 100,
        },
    ])

    return {
        "summary": {
            "Total annual operating cost [€/year]": (
                cost_breakdown["Total annual operating cost [€/year]"].sum()
            )
        },
        "breakdown": cost_breakdown,
        "remaining_life": remaining_life,
    }

def print_energy_rows(r, rows):
    """Print selected annual energy flows in MWh."""
    for label, column in rows:
        print(f"{label + ':':32s}{r[column].sum() / 1000:10.2f} MWh")



def calculate_diesel_comparison(load, cf_wind, cf_pv, current_results):
    """Confronta il sistema corrente con il riferimento diesel-only.

    Il riferimento usa lo stesso profilo di carico, lo stesso modello del motore
    diesel e lo stesso modello dell'alternatore, ma disattiva PV, wind, batteria,
    elettrolizzatore, accumulo H2 e fuel cell.

    Il generatore di riferimento e' quello esistente da DIESEL_UNIT_SIZE kW.
    Le taglie correnti vengono ripristinate al termine del calcolo.
    """
    global PV_SIZE, N_WT, WIND_SIZE, BATTERY_SIZE
    global EL_SIZE, H2_STORAGE_SIZE, FC_SIZE, DIESEL_SIZE

    saved_sizes = {
        "PV_SIZE": PV_SIZE,
        "N_WT": N_WT,
        "WIND_SIZE": WIND_SIZE,
        "BATTERY_SIZE": BATTERY_SIZE,
        "EL_SIZE": EL_SIZE,
        "H2_STORAGE_SIZE": H2_STORAGE_SIZE,
        "FC_SIZE": FC_SIZE,
        "DIESEL_SIZE": DIESEL_SIZE,
    }

    try:
        # Riferimento: un solo genset esistente, senza altri componenti.
        PV_SIZE = 0.0
        N_WT = 0
        WIND_SIZE = 0.0
        BATTERY_SIZE = 0.0
        EL_SIZE = 0.0
        H2_STORAGE_SIZE = 0.0
        FC_SIZE = 0.0
        DIESEL_SIZE = float(DIESEL_UNIT_SIZE)

        diesel_only_results = simulate(load, cf_wind, cf_pv)

    finally:
        # Ripristina esattamente la configurazione che l'utente sta simulando.
        PV_SIZE = saved_sizes["PV_SIZE"]
        N_WT = saved_sizes["N_WT"]
        WIND_SIZE = saved_sizes["WIND_SIZE"]
        BATTERY_SIZE = saved_sizes["BATTERY_SIZE"]
        EL_SIZE = saved_sizes["EL_SIZE"]
        H2_STORAGE_SIZE = saved_sizes["H2_STORAGE_SIZE"]
        FC_SIZE = saved_sizes["FC_SIZE"]
        DIESEL_SIZE = saved_sizes["DIESEL_SIZE"]

    baseline_fuel = float(diesel_only_results["diesel_fuel"].sum())
    current_fuel = float(current_results["diesel_fuel"].sum())

    fuel_saved = baseline_fuel - current_fuel
    fuel_reduction_percent = (
        100.0 * fuel_saved / baseline_fuel
        if baseline_fuel > 0 else np.nan
    )

    baseline_lpsp = (
        100.0 * diesel_only_results["unserved"].sum()
        / diesel_only_results["load"].sum()
    )
    current_lpsp = (
        100.0 * current_results["unserved"].sum()
        / current_results["load"].sum()
    )

    baseline_co2 = baseline_fuel * DIESEL_EMISSION_FACTOR / 1000.0
    current_co2 = current_fuel * DIESEL_EMISSION_FACTOR / 1000.0
    co2_avoided = baseline_co2 - current_co2
    baseline_operating_time = int(
        np.count_nonzero(diesel_only_results["diesel_power"].to_numpy() > 1e-9)
    )
    current_operating_time = int(
        np.count_nonzero(current_results["diesel_power"].to_numpy() > 1e-9)
    )
    baseline_diesel_energy_kwh = float(diesel_only_results["diesel_power"].sum())
    current_diesel_energy_kwh = float(current_results["diesel_power"].sum())

    return {
        "diesel_only_fuel_l": baseline_fuel,
        "current_fuel_l": current_fuel,
        "fuel_saved_l": fuel_saved,
        "fuel_reduction_percent": fuel_reduction_percent,
        "diesel_only_lpsp_percent": baseline_lpsp,
        "current_lpsp_percent": current_lpsp,
        "diesel_only_co2_t": baseline_co2,
        "current_co2_t": current_co2,
        "co2_avoided_t": co2_avoided,
        "diesel_only_operating_hours": baseline_operating_time,
        "current_diesel_operating_hours": current_operating_time,
        "diesel_only_generation_kwh": baseline_diesel_energy_kwh,
        "current_diesel_generation_kwh": current_diesel_energy_kwh,
    }


def calculate_annual_cash_opex(r, metrics=None):
    """Annual cash OPEX for payback; replacements/degradation reserves excluded."""
    metrics = metrics or calculate_operational_metrics(r)
    diesel_size = r.attrs["diesel_size"]
    el_capex = electrolyzer_specific_capex(EL_SIZE) * EL_SIZE
    fc_capex = fuel_cell_specific_capex(FC_SIZE) * FC_SIZE
    h2_mass = H2_STORAGE_SIZE * 3.6 / H2_LHV if H2_STORAGE_SIZE > 0 else 0.0
    h2_throughput = metrics["h2_produced_energy"] + metrics["h2_consumed_energy"]

    return {
        "Wind": (
            WIND_FIXED_OM * WIND_SIZE
            + WIND_VARIABLE_OM * r["wind"].sum()
        ),
        "PV": (
            PV_FIXED_OM * PV_SIZE
            + PV_VARIABLE_OM * r["pv"].sum()
        ),
        "Battery Li-ion": (
            BATTERY_FIXED_OM * BATTERY_SIZE
            + BATTERY_VARIABLE_OM * metrics["battery_throughput"]
        ),
        "PEM electrolyzer": (
            EL_TOTAL_OM_FRACTION / 3 * el_capex
            + (2 / 3) * EL_TOTAL_OM_FRACTION
            * el_capex / 8760 * metrics["el_operating_time"]
            + EL_FIXED_OM_PER_KW_YEAR * EL_SIZE
        ),
        "PEM fuel cell": (
            FC_TOTAL_OM_FRACTION / 3 * fc_capex
            + (2 / 3) * FC_TOTAL_OM_FRACTION
            * fc_capex / 8760 * metrics["fc_operating_time"]
            + FC_FIXED_OM_PER_KW_YEAR * FC_SIZE
            + FC_OM_PER_OPERATING_HOUR * metrics["fc_operating_time"]
        ),
        "H2 tank": (
            H2_TANK_FIXED_OM_FRACTION * H2_TANK_CAPEX * h2_mass
            + H2_TANK_FIXED_OM_PER_KG_YEAR * h2_mass
        ),
        "H2 auxiliaries": H2_SYSTEM_VARIABLE_OM * h2_throughput,
        "Diesel generator": (
            DIESEL_FIXED_OM_FRACTION * DIESEL_CAPEX * diesel_size
            + DIESEL_VARIABLE_OM * r["diesel_power"].sum()
            + DIESEL_OM_PER_OPERATING_HOUR * metrics["diesel_operating_time"]
            + DIESEL_OM_PER_L * r["diesel_fuel"].sum()
            + DIESEL_FUEL_COST * r["diesel_fuel"].sum()
        ),
    }


def calculate_payback_analysis(r, diesel_comparison, metrics=None):
    """Simple payback versus the existing diesel-only reference."""
    metrics = metrics or calculate_operational_metrics(r)
    current_opex = calculate_annual_cash_opex(r, metrics)
    current_total_opex = float(sum(current_opex.values()))

    baseline_fuel_l = diesel_comparison["diesel_only_fuel_l"]
    baseline_fuel_cost = baseline_fuel_l * DIESEL_FUEL_COST
    baseline_diesel_opex = (
        baseline_fuel_cost
        + DIESEL_VARIABLE_OM * diesel_comparison["diesel_only_generation_kwh"]
        + DIESEL_OM_PER_OPERATING_HOUR
        * diesel_comparison["diesel_only_operating_hours"]
        + DIESEL_OM_PER_L * baseline_fuel_l
        + DIESEL_FIXED_OM_FRACTION * DIESEL_CAPEX * DIESEL_UNIT_SIZE
    )

    gross_fuel_saving = diesel_comparison["fuel_saved_l"] * DIESEL_FUEL_COST
    net_annual_saving = baseline_diesel_opex - current_total_opex

    h2_mass = H2_STORAGE_SIZE * 3.6 / H2_LHV if H2_STORAGE_SIZE > 0 else 0.0
    capex = {
        "PV": PV_CAPEX * PV_SIZE,
        "Wind": WIND_CAPEX * WIND_SIZE,
        "Battery Li-ion": BATTERY_CAPEX * BATTERY_SIZE,
        "PEM electrolyzer": electrolyzer_specific_capex(EL_SIZE) * EL_SIZE,
        "H2 tank": H2_TANK_CAPEX * h2_mass,
        "H2 auxiliaries": h2_auxiliary_capex(),
        "PEM fuel cell": fuel_cell_specific_capex(FC_SIZE) * FC_SIZE,
    }
    incremental_capex = float(sum(capex.values()))
    simple_pbt = (
        incremental_capex / net_annual_saving
        if net_annual_saving > 0 else np.inf
    )

    current_lpsp = float(diesel_comparison["current_lpsp_percent"])
    pbt_comparable = current_lpsp <= 1e-6

    summary = {
        "Diesel fuel saved [L/year]": diesel_comparison["fuel_saved_l"],
        "Gross diesel fuel-cost saving [EUR/year]": gross_fuel_saving,
        "Diesel-only annual cash OPEX [EUR/year]": baseline_diesel_opex,
        "Current-system annual cash OPEX [EUR/year]": current_total_opex,
        "Net annual operating saving [EUR/year]": net_annual_saving,
        "Incremental system CAPEX [EUR]": incremental_capex,
        "Simple payback time [years]": simple_pbt,
        "PBT reliability comparable [-]": pbt_comparable,
        "Current system LPSP [%]": current_lpsp,
    }

    summary_table = pd.DataFrame([
        ("Diesel fuel saved", diesel_comparison["fuel_saved_l"], "L/year"),
        ("Gross diesel fuel-cost saving", gross_fuel_saving, "EUR/year"),
        ("Diesel-only annual cash OPEX", baseline_diesel_opex, "EUR/year"),
        ("Current-system annual cash OPEX", current_total_opex, "EUR/year"),
        ("Net annual operating saving", net_annual_saving, "EUR/year"),
        ("Incremental system CAPEX", incremental_capex, "EUR"),
        ("Simple payback time", simple_pbt, "years"),
        ("Current system LPSP", current_lpsp, "%"),
        ("PBT reliability comparable", "YES" if pbt_comparable else "NO", "-"),
    ], columns=["Indicator", "Value", "Unit"])

    capex_breakdown = pd.DataFrame(
        [(component, value) for component, value in capex.items()],
        columns=["Component", "Incremental CAPEX [EUR]"]
    )

    opex_rows = [("Diesel-only", "Diesel generator", baseline_diesel_opex)]
    opex_rows.extend(
        ("Current system", component, value)
        for component, value in current_opex.items()
    )
    opex_breakdown = pd.DataFrame(
        opex_rows,
        columns=["Scenario", "Component", "Annual cash OPEX [EUR/year]"]
    )

    return {
        "summary": summary,
        "summary_table": summary_table,
        "capex_breakdown": capex_breakdown,
        "opex_breakdown": opex_breakdown,
    }

def print_payback_analysis(payback):
    """Print savings and simple payback versus diesel-only."""
    s = payback["summary"]
    print("\nPAYBACK ANALYSIS VS DIESEL-ONLY")
    print(f"{'Diesel fuel saved:':42s}{s['Diesel fuel saved [L/year]']:14.2f} L/year")
    print(f"{'Gross fuel-cost saving:':42s}{s['Gross diesel fuel-cost saving [EUR/year]']:14.2f} EUR/year")
    print(f"{'Diesel-only annual cash OPEX:':42s}{s['Diesel-only annual cash OPEX [EUR/year]']:14.2f} EUR/year")
    print(f"{'Current-system annual cash OPEX:':42s}{s['Current-system annual cash OPEX [EUR/year]']:14.2f} EUR/year")
    print(f"{'Net annual operating saving:':42s}{s['Net annual operating saving [EUR/year]']:14.2f} EUR/year")
    print(f"{'Incremental system CAPEX:':42s}{s['Incremental system CAPEX [EUR]']:14.2f} EUR")
    if np.isfinite(s["Simple payback time [years]"]):
        print(f"{'Simple payback time (PBT):':42s}{s['Simple payback time [years]']:14.2f} years")
    else:
        print(f"{'Simple payback time (PBT):':42s}{'NO PAYBACK':>14s}")
    print(f"{'Current-system LPSP:':42s}{s['Current system LPSP [%]']:14.4f} %")
    if not s["PBT reliability comparable [-]"]:
        print("  WARNING: this PBT is NOT directly comparable with diesel-only because")
        print("  the current configuration has unserved load (LPSP > 0).")
    print("  PBT = incremental system CAPEX / net annual cash saving.")
    print("  Existing diesel CAPEX is excluded because the reference uses the existing genset.")
    print("  Future replacements are handled in NPC/LCOE, not in this simple PBT.")


def print_diesel_comparison(comparison):
    """Stampa il confronto tra configurazione corrente e riferimento diesel-only."""
    print("\nDIESEL SAVING ANALYSIS")
    print(
        f"{'Diesel-only reference consumption:':38s}"
        f"{comparison['diesel_only_fuel_l']:12.2f} L/year"
    )
    print(
        f"{'Current system consumption:':38s}"
        f"{comparison['current_fuel_l']:12.2f} L/year"
    )
    print(
        f"{'Diesel fuel saved:':38s}"
        f"{comparison['fuel_saved_l']:12.2f} L/year"
    )
    print(
        f"{'Diesel fuel reduction:':38s}"
        f"{comparison['fuel_reduction_percent']:12.2f} %"
    )
    print(
        f"{'CO2 avoided:':38s}"
        f"{comparison['co2_avoided_t']:12.2f} tCO2/year"
    )
    print(
        f"{'Diesel-only reference LPSP:':38s}"
        f"{comparison['diesel_only_lpsp_percent']:12.4f} %"
    )
    print(
        f"{'Current system LPSP:':38s}"
        f"{comparison['current_lpsp_percent']:12.4f} %"
    )

    if comparison["current_lpsp_percent"] > 1e-9:
        print(
            "  NOTE: fuel reduction must be interpreted together with LPSP, "
            "because part of the load is not served in the current configuration."
        )


def print_summary(r, diesel_comparison=None):
    """Print annual energy balances and reliability indicators in English."""
    print("\nANNUAL ENERGY SUMMARY")
    print_energy_rows(r, [
        ("PV generation", "pv"),
        ("Wind generation", "wind"),
        ("Total renewable generation", "res_total"),
        ("Total electrical load", "load"),
    ])

    print("\nRENEWABLE ENERGY ALLOCATION")
    print_energy_rows(r, [
        ("Renewables to battery", "res_to_battery"),
        ("Renewables to electrolyzer", "res_to_el"),
        ("Renewables to load", "res_to_load"),
        ("Renewable curtailment", "curtailment"),
    ])

    print("\nLOAD SUPPLY ALLOCATION")
    print_energy_rows(r, [
        ("Renewables to load", "res_to_load"),
        ("Battery to load", "battery_to_load"),
        ("Fuel cell to load", "fc_to_load"),
        ("Diesel generator to load", "diesel_to_load"),
        ("Unserved energy", "unserved"),
    ])

    print("\nMINIMUM-LOAD EXCESS ALLOCATION")
    print_energy_rows(r, [
        ("Dispatchable generation to battery", "dispatch_to_battery"),
        ("Dispatchable generation curtailed", "dispatch_curtailment"),
    ])

    h2_metrics = calculate_operational_metrics(r)
    print("\nHYDROGEN SYSTEM PERFORMANCE")
    if np.isfinite(h2_metrics["el_average_load_fraction"]):
        print(
            f"{'Electrolyzer average load fraction:':38s}"
            f"{100 * h2_metrics['el_average_load_fraction']:12.2f} %"
        )
        print(
            f"{'Electrolyzer average efficiency:':38s}"
            f"{100 * h2_metrics['el_average_efficiency']:12.2f} %"
        )
        print(
            f"{'H2 produced (LHV):':38s}"
            f"{h2_metrics['h2_produced_energy'] / 1000:12.2f} MWh_H2/year"
        )
    else:
        print(f"{'Electrolyzer:':38s}{'not operated':>12s}")

    if np.isfinite(h2_metrics["fc_average_load_fraction"]):
        print(
            f"{'Fuel cell average load fraction:':38s}"
            f"{100 * h2_metrics['fc_average_load_fraction']:12.2f} %"
        )
        print(
            f"{'Fuel cell average efficiency:':38s}"
            f"{100 * h2_metrics['fc_average_efficiency']:12.2f} %"
        )
        print(
            f"{'H2 consumed (LHV):':38s}"
            f"{h2_metrics['h2_consumed_energy'] / 1000:12.2f} MWh_H2/year"
        )
    else:
        print(f"{'Fuel cell:':38s}{'not operated':>12s}")

    if np.isfinite(h2_metrics["h2_conversion_chain_efficiency"]):
        print(
            f"{'EL x FC conversion-chain efficiency:':38s}"
            f"{100 * h2_metrics['h2_conversion_chain_efficiency']:12.2f} %"
        )

    print("\nDIESEL ENGINE AND ALTERNATOR")
    print_energy_rows(r, [
        ("Diesel electrical output", "diesel_power"),
        ("Diesel engine shaft energy", "diesel_shaft_power"),
        ("Alternator losses", "alternator_loss"),
    ])
    diesel_electric = r["diesel_power"].sum()
    if diesel_electric > 0:
        mean_eta = diesel_electric / r["diesel_shaft_power"].sum()
        print(f"{'Mean alternator efficiency:':32s}{100 * mean_eta:10.2f} %")

    lpsp = 100 * r["unserved"].sum() / r["load"].sum()
    fuel = r["diesel_fuel"].sum()  # L
    co2 = fuel * DIESEL_EMISSION_FACTOR / 1000  # tCO2
    print("\nRELIABILITY AND DIESEL USE")
    print(f"{'LPSP:':38s}{lpsp:12.3f} %")
    print(f"{'Diesel fuel consumption:':38s}{fuel:12.2f} L/year")
    print(f"{'Diesel CO2 emissions:':38s}{co2:12.2f} tCO2/year")

    if diesel_comparison is not None:
        print(
            f"{'Diesel-only reference consumption:':38s}"
            f"{diesel_comparison['diesel_only_fuel_l']:12.2f} L/year"
        )
        print(
            f"{'Diesel fuel saved:':38s}"
            f"{diesel_comparison['fuel_saved_l']:12.2f} L/year"
        )
        print(
            f"{'Diesel fuel reduction:':38s}"
            f"{diesel_comparison['fuel_reduction_percent']:12.2f} %"
        )
        print(
            f"{'CO2 avoided:':38s}"
            f"{diesel_comparison['co2_avoided_t']:12.2f} tCO2/year"
        )
        print(
            f"{'Diesel-only reference LPSP:':38s}"
            f"{diesel_comparison['diesel_only_lpsp_percent']:12.4f} %"
        )

        if diesel_comparison["current_lpsp_percent"] > 1e-9:
            print(
                "  NOTE: fuel reduction must be interpreted together with LPSP, "
                "because part of the load is not served."
            )


def print_economic_summary(economics):
    """Stampa i principali risultati economici e le vite calcolate."""
    summary = economics["summary"]
    kpis = economics["kpis"]
    print("\nECONOMIC SUMMARY")
    print(f"{'Total CAPEX:':42s}{summary['Total CAPEX [M€]']:14.3f} M€")
    print(f"{'OPEX present value:':42s}{summary['Total OPEX present value [M€]']:14.3f} M€")
    print(f"{'Replacement present value:':42s}{summary['Total replacement present value [M€]']:14.3f} M€")
    print(f"{'Salvage present value:':42s}{summary['Total salvage present value [M€]']:14.3f} M€")
    print(f"{'Total NPC:':42s}{summary['Total NPC [M€]']:14.3f} M€")
    print(f"{'LCOE:':42s}{summary['LCOE [€/kWh]']:14.4f} €/kWh")

    print("\nCALCULATED COMPONENT LIFETIMES")
    for label in [
        "Battery lifetime [years]",
        "Electrolyzer stack lifetime [years]",
        "Fuel cell stack lifetime [years]",
        "Diesel lifetime [years]",
    ]:
        print(f"{label + ':':42s}{kpis[label]:14.2f} years")


def print_annual_operation(annual_operation):
    """Stampa costi operativi annuali e vita residua a fine anno."""
    print("\nANNUAL OPERATING COSTS")
    for _, row in annual_operation["breakdown"].iterrows():
        print(
            f"{row['Component'] + ':':30s}"
            f"{row['Total annual operating cost [€/year]']:14.2f} €/year"
        )
    total = annual_operation["summary"][
        "Total annual operating cost [€/year]"
    ]
    print(f"{'Total:':30s}{total:14.2f} €/year")

    print("\nEND-OF-YEAR REMAINING LIFE")
    for _, row in annual_operation["remaining_life"].iterrows():
        print(
            f"{row['Component'] + ':':30s}"
            f"{row['End-of-year remaining life [%]']:14.2f} %"
        )


def export_results_to_excel(r, economics=None, annual_operation=None,
                            diesel_comparison=None, payback_analysis=None,
                            file_name="simulation_results.xlsx"):
    """Save a formatted Excel workbook with annual and hourly results."""
    lpsp = 100 * r["unserved"].sum() / r["load"].sum()
    fuel = r["diesel_fuel"].sum()  # L
    co2 = fuel * DIESEL_EMISSION_FACTOR / 1000  # tCO2
    mwh = lambda column: r[column].sum() / 1000

    summary_rows = [
        ["ANNUAL ENERGY SUMMARY", None, None],
        ["Indicator", "Value", "Unit"],
        ["PV generation", mwh("pv"), "MWh"],
        ["Wind generation", mwh("wind"), "MWh"],
        ["Total renewable generation", mwh("res_total"), "MWh"],
        ["Total electrical load", mwh("load"), "MWh"],
        [None, None, None],
        ["RENEWABLE ENERGY ALLOCATION", None, None],
        ["Indicator", "Value", "Unit"],
        ["Renewables to battery", mwh("res_to_battery"), "MWh"],
        ["Renewables to electrolyzer", mwh("res_to_el"), "MWh"],
        ["Renewables to load", mwh("res_to_load"), "MWh"],
        ["Renewable curtailment", mwh("curtailment"), "MWh"],
        [None, None, None],
        ["LOAD SUPPLY ALLOCATION", None, None],
        ["Indicator", "Value", "Unit"],
        ["Renewables to load", mwh("res_to_load"), "MWh"],
        ["Battery to load", mwh("battery_to_load"), "MWh"],
        ["Fuel cell to load", mwh("fc_to_load"), "MWh"],
        ["Diesel generator to load", mwh("diesel_to_load"), "MWh"],
        ["Unserved energy", mwh("unserved"), "MWh"],
        [None, None, None],
        ["MINIMUM-LOAD EXCESS ALLOCATION", None, None],
        ["Indicator", "Value", "Unit"],
        ["Dispatchable generation to battery", mwh("dispatch_to_battery"), "MWh"],
        ["Dispatchable generation curtailed", mwh("dispatch_curtailment"), "MWh"],
        [None, None, None],
        ["RELIABILITY AND DIESEL USE", None, None],
        ["Indicator", "Value", "Unit"],
        ["LPSP", lpsp, "%"],
        ["Diesel fuel consumption", fuel, "L"],
        ["Diesel CO2 emissions", co2, "tCO2"],
        [None, None, None],
        ["DIESEL ENGINE AND ALTERNATOR", None, None],
        ["Indicator", "Value", "Unit"],
        ["Diesel electrical output", mwh("diesel_power"), "MWh"],
        ["Diesel engine shaft energy", mwh("diesel_shaft_power"), "MWh"],
        ["Alternator losses", mwh("alternator_loss"), "MWh"],
        [None, None, None],
        ["HYDROGEN SYSTEM PERFORMANCE", None, None],
        ["Indicator", "Value", "Unit"],
        ["Electrolyzer average load fraction",
         100 * calculate_operational_metrics(r)["el_average_load_fraction"], "%"],
        ["Electrolyzer average system efficiency",
         100 * calculate_operational_metrics(r)["el_average_efficiency"], "%"],
        ["H2 produced (LHV)", r["h2_produced"].sum() / 1000, "MWh_H2"],
        ["Fuel cell average load fraction",
         100 * calculate_operational_metrics(r)["fc_average_load_fraction"], "%"],
        ["Fuel cell average system efficiency",
         100 * calculate_operational_metrics(r)["fc_average_efficiency"], "%"],
        ["H2 consumed (LHV)", r["h2_consumed"].sum() / 1000, "MWh_H2"],
        ["EL x FC conversion-chain efficiency",
         100 * calculate_operational_metrics(r)["h2_conversion_chain_efficiency"], "%"],
    ]
    summary = pd.DataFrame(summary_rows)
    hourly = r.rename(columns={
        "load": "Electrical load [kW]", "pv": "PV generation [kW]",
        "wind": "Wind generation [kW]", "res_total": "Total renewable generation [kW]",
        "res_to_load": "Renewables to load [kW]", "battery_to_load": "Battery to load [kW]",
        "fc_to_load": "Fuel cell to load [kW]", "diesel_to_load": "Diesel generator to load [kW]",
        "unserved": "Unserved energy [kW]", "res_to_battery": "Renewables to battery [kW]",
        "dispatch_to_battery": "Dispatchable generation to battery [kW]",
        "res_to_el": "Renewables to electrolyzer [kW]", "curtailment": "Renewable curtailment [kW]",
        "dispatch_curtailment": "Dispatchable generation curtailment [kW]",
        "battery_power": "Battery power at busbar [kW]", "fc_power": "Fuel cell power [kW]",
        "el_power": "Electrolyzer power [kW]",
        "el_efficiency": "Electrolyzer system efficiency [-]",
        "fc_efficiency": "Fuel cell system efficiency [-]",
        "h2_produced": "H2 produced [kWh_H2 LHV]",
        "h2_consumed": "H2 consumed [kWh_H2 LHV]",
        "diesel_power": "Diesel generator power [kW]",
        "diesel_fuel": "Diesel fuel consumption [L]", "battery_soc": "Battery SOC [-]",
        "h2_soc": "Hydrogen storage SOC [-]",
        "diesel_shaft_power": "Diesel engine shaft power [kW]",
        "alternator_loss": "Alternator losses [kW]",
    })

    with pd.ExcelWriter(file_name, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Annual summary", index=False, header=False)
        hourly.to_excel(writer, sheet_name="Hourly results", index_label="Hour")

        economic_inputs = pd.DataFrame([
            ("Cost scenario", COST_SCENARIO, "-"),
            ("Cost scenario key", COST_SCENARIO_KEY, "-"),
            ("Project lifetime", PROJECT_LIFETIME, "year"),
            ("Real discount rate", 100 * DISCOUNT_RATE, "%"),
            ("USD to EUR", USD_TO_EUR, "EUR/USD"),
            ("CAD to EUR", CAD_TO_EUR, "EUR/CAD"),
            ("PV CAPEX", PV_CAPEX, "EUR/kW"),
            ("PV fixed O&M", PV_FIXED_OM, "EUR/(kW year)"),
            ("PV variable O&M", PV_VARIABLE_OM, "EUR/kWh generated"),
            ("Wind CAPEX", WIND_CAPEX, "EUR/kW"),
            ("Wind fixed O&M", WIND_FIXED_OM, "EUR/(kW year)"),
            ("Wind variable O&M", WIND_VARIABLE_OM, "EUR/kWh generated"),
            ("Battery CAPEX", BATTERY_CAPEX, "EUR/kWh"),
            ("Battery replacement", BATTERY_REPLACEMENT_COST, "EUR/kWh"),
            ("Battery fixed O&M", BATTERY_FIXED_OM, "EUR/(kWh year)"),
            ("Battery variable O&M", BATTERY_VARIABLE_OM, "EUR/kWh throughput"),
            ("Electrolyzer CAPEX input/reference", EL_CAPEX, "EUR/kW"),
            ("Electrolyzer effective CAPEX", electrolyzer_specific_capex(EL_SIZE), "EUR/kW"),
            ("Electrolyzer size scaling active", "YES" if EL_USE_SIZE_SCALING else "NO", "-"),
            ("Electrolyzer stack replacement", 100 * EL_STACK_REPLACEMENT_FRACTION, "% CAPEX"),
            ("Electrolyzer total fractional O&M", 100 * EL_TOTAL_OM_FRACTION, "% CAPEX/year"),
            ("Electrolyzer fixed O&M", EL_FIXED_OM_PER_KW_YEAR, "EUR/(kW year)"),
            ("H2 tank CAPEX", H2_TANK_CAPEX, "EUR/kg H2"),
            ("H2 tank replacement", H2_TANK_REPLACEMENT_COST, "EUR/kg H2"),
            ("H2 tank fixed O&M fraction", 100 * H2_TANK_FIXED_OM_FRACTION, "% CAPEX/year"),
            ("H2 tank fixed O&M absolute", H2_TANK_FIXED_OM_PER_KG_YEAR, "EUR/(kg year)"),
            ("Water treatment CAPEX", H2_WATER_TREATMENT_CAPEX_PER_KW_EL, "EUR/kW_EL"),
            ("Compressor CAPEX", H2_COMPRESSOR_CAPEX_PER_KW_EL, "EUR/kW_EL"),
            ("H2 system variable O&M", H2_SYSTEM_VARIABLE_OM, "EUR/kWh_H2 throughput"),
            ("Fuel cell CAPEX input/reference", FC_CAPEX, "EUR/kW"),
            ("Fuel cell effective CAPEX", fuel_cell_specific_capex(FC_SIZE), "EUR/kW"),
            ("Fuel cell size scaling active", "YES" if FC_USE_SIZE_SCALING else "NO", "-"),
            ("Fuel cell stack replacement", 100 * FC_STACK_REPLACEMENT_FRACTION, "% CAPEX"),
            ("Fuel cell total fractional O&M", 100 * FC_TOTAL_OM_FRACTION, "% CAPEX/year"),
            ("Fuel cell fixed O&M", FC_FIXED_OM_PER_KW_YEAR, "EUR/(kW year)"),
            ("Fuel cell operating O&M", FC_OM_PER_OPERATING_HOUR, "EUR/operating hour"),
            ("Diesel initial CAPEX", DIESEL_CAPEX, "EUR/kW"),
            ("Diesel replacement", DIESEL_REPLACEMENT_COST, "EUR/kW"),
            ("Diesel variable O&M", DIESEL_VARIABLE_OM, "EUR/kWh_el"),
            ("Diesel operating O&M", DIESEL_OM_PER_OPERATING_HOUR, "EUR/h"),
            ("Diesel fuel-linked O&M", DIESEL_OM_PER_L, "EUR/L"),
            ("Diesel fuel price", DIESEL_FUEL_COST, "EUR/L"),
        ], columns=["Economic input", "Value", "Unit"])
        economic_inputs.to_excel(writer, sheet_name="Economic inputs", index=False)

        if diesel_comparison is not None:
            diesel_comparison_table = pd.DataFrame([
                ("Diesel-only reference consumption",
                 diesel_comparison["diesel_only_fuel_l"], "L/year"),
                ("Current system consumption",
                 diesel_comparison["current_fuel_l"], "L/year"),
                ("Diesel fuel saved",
                 diesel_comparison["fuel_saved_l"], "L/year"),
                ("Diesel fuel reduction",
                 diesel_comparison["fuel_reduction_percent"], "%"),
                ("Diesel-only reference LPSP",
                 diesel_comparison["diesel_only_lpsp_percent"], "%"),
                ("Current system LPSP",
                 diesel_comparison["current_lpsp_percent"], "%"),
                ("Diesel-only reference CO2",
                 diesel_comparison["diesel_only_co2_t"], "tCO2/year"),
                ("Current system CO2",
                 diesel_comparison["current_co2_t"], "tCO2/year"),
                ("CO2 avoided",
                 diesel_comparison["co2_avoided_t"], "tCO2/year"),
            ], columns=["Indicator", "Value", "Unit"])
            diesel_comparison_table.to_excel(
                writer, sheet_name="Diesel comparison", index=False
            )

        if payback_analysis is not None:
            payback_sheet = "Payback analysis"
            payback_analysis["summary_table"].to_excel(
                writer, sheet_name=payback_sheet, index=False, startrow=0
            )
            capex_start = len(payback_analysis["summary_table"]) + 3
            payback_analysis["capex_breakdown"].to_excel(
                writer, sheet_name=payback_sheet, index=False, startrow=capex_start
            )
            opex_start = capex_start + len(payback_analysis["capex_breakdown"]) + 3
            payback_analysis["opex_breakdown"].to_excel(
                writer, sheet_name=payback_sheet, index=False, startrow=opex_start
            )

        if economics is not None:
            economic_summary = pd.DataFrame(
                list(economics["summary"].items()),
                columns=["Indicator", "Value"]
            )
            operational_kpis = pd.DataFrame(
                list(economics["kpis"].items()),
                columns=["Indicator", "Value"]
            )
            economic_summary.to_excel(
                writer, sheet_name="Economic summary", index=False
            )
            economics["breakdown"].to_excel(
                writer, sheet_name="Economic breakdown", index=False
            )
            operational_kpis.to_excel(
                writer, sheet_name="Operational KPIs", index=False
            )
        if annual_operation is not None:
            total_row = {
                "Component": "TOTAL",
                "Total annual operating cost [€/year]": (
                    annual_operation["summary"][
                        "Total annual operating cost [€/year]"
                    ]
                ),
            }
            annual_cost_export = pd.concat(
                [annual_operation["breakdown"], pd.DataFrame([total_row])],
                ignore_index=True,
            )
            annual_cost_export.to_excel(
                writer, sheet_name="Annual operating costs", index=False
            )
            annual_operation["remaining_life"].to_excel(
                writer, sheet_name="Remaining life", index=False
            )

        workbook = writer.book
        summary_sheet = writer.sheets["Annual summary"]
        hourly_sheet = writer.sheets["Hourly results"]
        dark_blue = "17365D"
        light_blue = "D9EAF7"
        white_font = Font(color="FFFFFF", bold=True, size=12)
        section_fill = PatternFill("solid", fgColor=dark_blue)
        header_fill = PatternFill("solid", fgColor=light_blue)
        thin_blue = Side(style="thin", color="A6A6A6")

        summary_sheet.sheet_view.showGridLines = True
        summary_sheet.column_dimensions["A"].width = 34
        summary_sheet.column_dimensions["B"].width = 16
        summary_sheet.column_dimensions["C"].width = 12
        for row_number in [1, 8, 15, 23, 28, 34, 40]:
            summary_sheet.merge_cells(start_row=row_number, start_column=1, end_row=row_number, end_column=3)
            cell = summary_sheet.cell(row=row_number, column=1)
            cell.fill = section_fill
            cell.font = white_font
            cell.alignment = Alignment(horizontal="left")
            summary_sheet.row_dimensions[row_number].height = 24
        for row_number in [2, 9, 16, 24, 29, 35, 41]:
            for cell in summary_sheet[row_number]:
                cell.fill = header_fill
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center")
        for row in summary_sheet.iter_rows(
            min_row=1, max_row=summary_sheet.max_row, min_col=1, max_col=3
        ):
            for cell in row:
                cell.border = Border(bottom=thin_blue)
                if cell.column == 2 and isinstance(cell.value, (int, float)):
                    cell.number_format = "0.00"
                if cell.column == 3:
                    cell.alignment = Alignment(horizontal="center")

        hourly_sheet.freeze_panes = "A2"
        hourly_sheet.auto_filter.ref = hourly_sheet.dimensions
        hourly_sheet.sheet_view.showGridLines = True
        for cell in hourly_sheet[1]:
            cell.fill = section_fill
            cell.font = Font(color="FFFFFF", bold=True)
            cell.alignment = Alignment(horizontal="center", wrap_text=True)
        hourly_sheet.row_dimensions[1].height = 35
        hourly_sheet.column_dimensions["A"].width = 10
        for column in range(2, hourly_sheet.max_column + 1):
            hourly_sheet.column_dimensions[hourly_sheet.cell(row=1, column=column).column_letter].width = 24
        for row in hourly_sheet.iter_rows(min_row=2, max_row=hourly_sheet.max_row, min_col=2, max_col=hourly_sheet.max_column):
            for cell in row:
                cell.number_format = "0.000"
        workbook.active = 0

        formatted_sheets = []
        if diesel_comparison is not None:
            formatted_sheets.append("Diesel comparison")
        if payback_analysis is not None:
            formatted_sheets.append("Payback analysis")
        if economics is not None:
            formatted_sheets.extend([
                "Economic summary", "Economic breakdown",
                "Operational KPIs",
            ])
        if annual_operation is not None:
            formatted_sheets.extend([
                "Annual operating costs", "Remaining life",
            ])

        for sheet_name in formatted_sheets:
            sheet = writer.sheets[sheet_name]
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions
            sheet.sheet_view.showGridLines = True
            for cell in sheet[1]:
                cell.fill = section_fill
                cell.font = Font(color="FFFFFF", bold=True)
                cell.alignment = Alignment(horizontal="center", wrap_text=True)
            sheet.row_dimensions[1].height = 32
            for column in range(1, sheet.max_column + 1):
                header = str(sheet.cell(row=1, column=column).value or "")
                sheet.column_dimensions[
                    sheet.cell(row=1, column=column).column_letter
                ].width = min(max(len(header) + 4, 16), 34)
            for row in sheet.iter_rows(
                min_row=2, max_row=sheet.max_row,
                min_col=1, max_col=sheet.max_column
            ):
                for cell in row:
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = "0.000"

    print(f"\nCreated the Excel workbook: {file_name}")


def plot_operation_lines(ax, period, title):
    """Plot the main power flows for a selected period."""
    for column, label, color in [
        ("load", "Electrical load", "load"),
        ("pv", "PV generation", "pv"),
        ("wind", "Wind generation", "wind"),
        ("battery_power", "Battery power (busbar)", "battery"),
        ("fc_power", "Fuel cell generation", "fuel_cell"),
        ("diesel_power", "Diesel generation", "diesel"),
    ]:
        linewidth = 2.8 if column == "load" else 1.8
        ax.plot(
            period.index, period[column], label=label,
            color=COLORS[color], linewidth=linewidth
        )
    ax.axhline(0, color="black", linewidth=0.8)
    ax.set(title=title, xlabel="Hour", ylabel="Power [kW]")
    ax.grid(alpha=0.3)
    ax.legend(ncol=3, fontsize=10)


def plot_period_results(period, period_name, x_label="Hour"):
    """Create separate line and area charts for one selected time period."""
    _, axes = plt.subplots(3, 1, figsize=(15, 13), constrained_layout=True)
    plot_operation_lines(axes[0], period, f"Power flows — {period_name}")
    axes[0].set_xlabel(x_label)

    # Stacked areas show how the electrical load is supplied.
    load_supply_columns = ["res_to_load", "battery_to_load", "fc_to_load", "diesel_to_load", "unserved"]
    load_supply_labels = ["Direct renewables", "Battery", "Fuel cell", "Diesel generator", "Unserved"]
    load_supply_colors = [COLORS["renewables"], COLORS["battery"], COLORS["fuel_cell"], COLORS["diesel"], COLORS["unserved"]]
    axes[1].stackplot(period.index, *[period[column] for column in load_supply_columns],
                      labels=load_supply_labels, colors=load_supply_colors, alpha=0.92)
    axes[1].plot(period.index, period["load"], color=COLORS["load"], linewidth=2.2, label="Electrical load")
    axes[1].set(title=f"Load supply allocation — {period_name}", xlabel=x_label, ylabel="Power [kW]")
    axes[1].grid(alpha=0.3)
    axes[1].legend(ncol=3, fontsize=10, loc="upper right")

    # Stacked areas show how the renewable generation is allocated.
    renewable_columns = ["res_to_load", "res_to_battery", "res_to_el", "curtailment"]
    renewable_labels = ["Direct load", "Battery charging", "Electrolyzer", "Curtailment"]
    renewable_colors = [COLORS["renewables"], COLORS["battery"], COLORS["electrolyzer"], COLORS["curtailment"]]
    axes[2].stackplot(period.index, *[period[column] for column in renewable_columns],
                      labels=renewable_labels, colors=renewable_colors, alpha=0.92)
    axes[2].plot(period.index, period["res_total"], color=COLORS["load"], linewidth=2.2, label="Total renewable generation")
    axes[2].set(title=f"Renewable energy allocation — {period_name}", xlabel=x_label, ylabel="Power [kW]")
    axes[2].grid(alpha=0.3)
    axes[2].legend(ncol=3, fontsize=10, loc="upper right")


def plot_economic_breakdown(economics):
    """One stacked cost bar with NPC and LCOE scales."""
    breakdown = economics["breakdown"]
    fig, npc_axis = plt.subplots(figsize=(8, 7), constrained_layout=True)
    lcoe_axis = npc_axis.twinx()
    npc_bottom = 0.0

    for _, row in breakdown.iterrows():
        component = row["Component"]
        color = COST_COLORS[component]
        npc_value = row["NPC [M€]"]
        npc_axis.bar(
            0, npc_value, bottom=npc_bottom, width=0.55,
            color=color, edgecolor="white", label=component
        )
        npc_bottom += npc_value

    total_lcoe = economics["summary"]["LCOE [€/kWh]"]

    npc_axis.set(
        title="NPC and LCOE by component",
        ylabel="NPC [M€]",
        xticks=[0],
        xticklabels=["System cost"],
        xlim=(-0.7, 0.7),
    )
    lcoe_axis.set_ylabel("LCOE [€/kWh]")
    npc_axis.set_ylim(0, npc_bottom * 1.12 if npc_bottom > 0 else 1)
    lcoe_axis.set_ylim(0, total_lcoe * 1.12 if total_lcoe > 0 else 1)
    npc_axis.text(
        0, npc_bottom * 1.02,
        f"{npc_bottom:.3g} M€  |  {total_lcoe:.3g} €/kWh",
        ha="center", va="bottom", fontweight="bold"
    )
    npc_axis.grid(axis="y", alpha=0.3)
    npc_axis.legend(loc="upper center", bbox_to_anchor=(0.5, -0.08), ncol=4)


def plot_annual_overview(r):
    """Full-year power-flow view using all original hourly values.

    No daily averaging or temporal aggregation is applied: every one of the
    8760 hourly simulation points is represented. The x-axis is expressed as
    day of year only to keep the annual graph readable.
    """
    hours = np.arange(len(r), dtype=float)
    days = hours / 24.0 + 1.0

    fig, ax = plt.subplots(figsize=(16, 6), constrained_layout=True)
    ax.plot(days, r["load"].to_numpy(), label="Electrical load",
            color=COLORS["load"], linewidth=1.5)
    ax.plot(days, r["res_total"].to_numpy(), label="Renewable generation",
            color=COLORS["renewables"], linewidth=1.1)
    ax.plot(days, r["diesel_power"].to_numpy(), label="Diesel generation",
            color=COLORS["diesel"], linewidth=1.0)
    ax.plot(days, r["fc_power"].to_numpy(), label="Fuel cell generation",
            color=COLORS["fuel_cell"], linewidth=1.0)
    ax.plot(days, r["el_power"].to_numpy(), label="Electrolyzer consumption",
            color=COLORS["electrolyzer"], linewidth=1.0)
    ax.set(
        title="Annual power-flow overview — full hourly resolution",
        xlabel="Day of year",
        ylabel="Power [kW]",
        xlim=(1, max(1.0, len(r) / 24.0)),
    )
    ax.grid(alpha=0.3)
    ax.legend(ncol=3, fontsize=10)


def plot_results(r, economics, annual=False, week_number=None, month_number=None):
    """Plot only the views selected by the user after the simulation."""
    if not MATPLOTLIB_AVAILABLE:
        print("Grafici non creati: installare matplotlib con pip install matplotlib")
        return

    if week_number is not None and not 1 <= week_number <= 52:
        raise ValueError("Week number must be between 1 and 52.")
    if month_number is not None and not 1 <= month_number <= 12:
        raise ValueError("Month number must be between 1 and 12.")

    rcParams.update({"font.size": 13, "axes.titlesize": 16, "axes.labelsize": 14,
                     "xtick.labelsize": 11, "ytick.labelsize": 11, "legend.fontsize": 10})

    # Annual auxiliary plots are created first. The detailed three-panel
    # whole-year dispatch plot is deliberately created LAST below.
    if annual:
        # State of charge over the complete year remains at hourly resolution.
        fig, ax = plt.subplots(figsize=(12, 4), constrained_layout=True)
        ax.plot(r["battery_soc"], label="Battery state of charge", color=COLORS["battery"], linewidth=2)
        ax.plot(r["h2_soc"], label="Hydrogen storage state of charge", color=COLORS["h2"], linewidth=2)
        ax.set(title="Annual storage state of charge", xlabel="Hour of year", ylabel="State of charge", ylim=(0, 1.05))
        ax.grid(alpha=0.3)
        ax.legend()

        # Annual energy-allocation pies.
        renewable_values = [r[column].sum() / 1000 for column in ["res_to_load", "res_to_battery", "res_to_el", "curtailment"]]
        load_values = [r[column].sum() / 1000 for column in ["res_to_load", "battery_to_load", "fc_to_load", "diesel_to_load", "unserved"]]
        fig, ax = plt.subplots(1, 2, figsize=(12, 5), constrained_layout=True)
        if sum(renewable_values) > 0:
            ax[0].pie(renewable_values, labels=["Direct load", "Battery", "Electrolyzer", "Curtailment"],
                      colors=[COLORS["renewables"], COLORS["battery"], COLORS["electrolyzer"], COLORS["curtailment"]], autopct="%1.1f%%", textprops={"fontsize": 11})
        else:
            ax[0].text(0.5, 0.5, "No renewable\ngeneration", ha="center", va="center")
            ax[0].axis("off")
        ax[0].set_title("Renewable energy allocation")

        if sum(load_values) > 0:
            ax[1].pie(load_values, labels=["Direct renewables", "Battery", "Fuel cell", "Diesel generator", "Unserved"],
                      colors=[COLORS["renewables"], COLORS["battery"], COLORS["fuel_cell"], COLORS["diesel"], COLORS["unserved"]], autopct="%1.1f%%", textprops={"fontsize": 11})
        else:
            ax[1].text(0.5, 0.5, "No load supply\ndata", ha="center", va="center")
            ax[1].axis("off")
        ax[1].set_title("Load supply allocation")

        plot_economic_breakdown(economics)

    # WEEK
    if week_number is not None:
        week_start = (week_number - 1) * 168
        week = r.iloc[week_start:week_start + 168]
        plot_period_results(week, f"Week {week_number}")

    # MONTH
    if month_number is not None:
        days_per_month = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
        month_start = sum(days_per_month[:month_number - 1]) * 24
        month = r.iloc[
            month_start:
            month_start + days_per_month[month_number - 1] * 24
        ]
        plot_period_results(month, f"Month {month_number}")

    # WHOLE YEAR — intentionally created last so it is the last figure window.
    if annual:
        plot_period_results(r, "Whole year", x_label="Hour of year")

    if annual or week_number is not None or month_number is not None:
        plt.show()


if __name__ == "__main__":
    APPLY_HORNSUND_CONSTRAINTS, SYSTEM_MODE = select_simulation_options()
    apply_system_mode(SYSTEM_MODE)

    # Two separate input screens: component sizes first, economic assumptions second.
    select_component_sizes(SYSTEM_MODE)
    select_economic_inputs()

    load, cf_wind, cf_pv = load_profiles(EXCEL_FILE, SHEET_NAME)
    print(f"Capacity factor medio PV:   {cf_pv.mean():.4f}")
    print(f"Capacity factor medio wind: {cf_wind.mean():.4f}")

    if APPLY_HORNSUND_CONSTRAINTS:
        design_constraints = calculate_design_constraints()
        print_design_constraint_check(design_constraints)
    else:
        design_constraints = None
        print("\nHORNSUND DESIGN CONSTRAINTS: DISABLED")

    results = simulate(load, cf_wind, cf_pv)

    # Annual storage sustainability constraints from the latest STORAGE_END model.
    battery_terminal_feasible = check_battery_terminal_soc(results, print_result=True)
    h2_terminal_feasible = check_h2_terminal_loh(results, print_result=True)
    storage_terminal_feasible = battery_terminal_feasible and h2_terminal_feasible

    diesel_comparison = calculate_diesel_comparison(
        load, cf_wind, cf_pv, results
    )
    print_summary(results, diesel_comparison)

    if APPLY_HORNSUND_CONSTRAINTS and DIESEL_SIZE > 0:
        fuel_constraint_feasible = print_fuel_constraint_check(results)
    else:
        fuel_constraint_feasible = True

    if SYSTEM_MODE in ("renewable_only", "battery_only_renewable") and results["unserved"].sum() > 1e-9:
        print(
            "\nWARNING: diesel is disabled, but the current renewable/storage sizes "
            "do not cover the full annual load. Check LPSP and unserved energy."
        )

    operational_metrics = calculate_operational_metrics(results)
    economics = calculate_economics(results, operational_metrics)
    annual_operation = calculate_annual_operation(results, operational_metrics)
    payback_analysis = calculate_payback_analysis(
        results, diesel_comparison, operational_metrics
    )
    print_economic_summary(economics)
    print_annual_operation(annual_operation)
    print_payback_analysis(payback_analysis)

    # Storage terminal conditions are active independently of optional Hornsund
    # site/design constraints.
    overall_feasible = storage_terminal_feasible
    if APPLY_HORNSUND_CONSTRAINTS:
        overall_feasible = (
            overall_feasible
            and design_constraints["overall_pre_feasible"]
            and fuel_constraint_feasible
        )

    print(
        f"\nOVERALL FEASIBILITY STATUS: "
        f"{'FEASIBLE' if overall_feasible else 'NOT FEASIBLE'}"
    )

    if SAVE_EXCEL_RESULTS:
        results_file = select_results_excel_file()
        if results_file is not None:
            export_results_to_excel(
                results,
                economics,
                annual_operation,
                diesel_comparison,
                payback_analysis,
                results_file,
            )
        else:
            print("\nSalvataggio risultati annullato: nessun file Excel creato.")
    if PLOT_RESULTS:
        plot_options = select_plot_options()
        if plot_options is not None:
            print(
                "\nPLOT SELECTION: "
                f"whole year={'YES' if plot_options['annual'] else 'NO'}, "
                f"week={plot_options['week'] if plot_options['week'] is not None else 'NO'}, "
                f"month={plot_options['month'] if plot_options['month'] is not None else 'NO'}"
            )
            plot_results(
                results,
                economics,
                annual=plot_options["annual"],
                week_number=plot_options["week"],
                month_number=plot_options["month"],
            )
        else:
            print("\nVisualizzazione grafici annullata.")
