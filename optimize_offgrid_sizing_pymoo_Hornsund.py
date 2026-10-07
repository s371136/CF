# -*- coding: utf-8 -*-

"""
Created on Wed Oct  7 16:16:20 2026

@author: 396373
"""

"""
Multi-objective Hornsund off-grid sizing with pymoo NSGA-II.

Questo ottimizzatore utilizza il simulatore rule-based di Hornsund selezionato
come modello orario interno e mantiene coerenza con le assunzioni e i vincoli
adottati per il dimensionamento del sistema energetico di Hornsund.

FUNZIONALITA PRINCIPALI:

1. Seleziona il simulatore rule-based tramite una finestra di selezione file.

2. Permette di attivare o disattivare i vincoli specifici di progetto/design
   di Hornsund.

3. Supporta quattro configurazioni di sistema:
   - sistema ibrido: batteria + H2 + diesel esistente;
   - sistema 100% rinnovabile: batteria + H2, senza diesel;
   - sistema ibrido battery-only: batteria + diesel esistente,
     senza catena dell'idrogeno;
   - sistema 100% rinnovabile battery-only: sola batteria,
     senza catena dell'idrogeno e senza diesel.

   Nei sistemi ibridi il diesel esistente viene mantenuto alla potenza
   installata e non rappresenta una variabile di dimensionamento.
   Nei sistemi 100% rinnovabili la potenza diesel viene fissata a zero.

4. Legge direttamente dal simulatore selezionato, quando disponibili,
   i principali parametri e vincoli di progetto, tra cui WT_RATED_POWER,
   N_WT_MAX, PV_MAX_SIZE e i parametri relativi agli accumuli e ai costi.

5. Mantiene coerenti N_WT e WIND_SIZE per ogni soluzione candidata,
   imponendo un numero intero di turbine.

6. Con i vincoli Hornsund attivi vengono applicati i limiti fisici e
   progettuali specifici del sito. Nella configurazione attuale la batteria
   prevista viene trattata come una scelta progettuale fissa da 500 kWh
   (FIX_HORNSUND_BATTERY_AT_PLANNED_SIZE = True).
   Se tale opzione viene disattivata, la batteria puo essere ottimizzata
   tra 0 e il limite di progetto di 500 kWh.

   Con i vincoli Hornsund disattivati, le taglie vengono invece ottimizzate
   su intervalli piu ampi; i relativi limiti superiori rappresentano
   esclusivamente limiti numerici del dominio di ricerca e non limiti
   fisici del sito.

7. Impone sempre i seguenti vincoli di fattibilita, indipendentemente
   dall'attivazione dei vincoli specifici di Hornsund:
   - LPSP <= MAX_LPSP_PERCENT;
   - SOC_batteria_finale >= SOC_batteria_iniziale;
   - LOH_H2_finale >= LOH_H2_iniziale.

8. Quando i vincoli Hornsund sono attivati, applica inoltre i vincoli
   specifici del sito, tra cui:
   - massima potenza PV installabile;
   - massimo numero di turbine;
   - vincolo sulla batteria prevista;
   - budget massimo del sistema H2;
   - limite sul consumo annuale di diesel;
   - eventuali ulteriori vincoli verificabili, ad esempio quelli logistici.

9. Calcola una volta il caso di riferimento diesel-only e determina la
   percentuale di decarbonizzazione di ciascuna soluzione rispetto alle
   emissioni di tale configurazione di riferimento.

10. Permette di utilizzare due metodi di ottimizzazione:
    - NSGA-II multi-obiettivo, per generare il fronte di Pareto tra
      prestazioni economiche e livello di decarbonizzazione;
    - epsilon-constraint, in cui per ciascun target minimo di
      decarbonizzazione viene eseguita una ottimizzazione single-objective
      con algoritmo genetico, minimizzando LCOE o NPC e imponendo
      Decarbonizzazione >= epsilon.

11. Nella modalita epsilon-constraint, l'utente puo scegliere direttamente
    dalla GUI il valore minimo, massimo e il passo dei target di
    decarbonizzazione.

12. La GUI permette di selezionare uno degli otto scenari economici:
    LOW, BASE, HIGH, quattro scenari derivati dalla letteratura e il caso
    sintetico MEDIAN.

    MEDIAN viene calcolato parametro per parametro utilizzando esclusivamente
    i quattro scenari di letteratura. LOW, BASE e HIGH sono scenari di
    sensitivita e sono quindi esclusi dal calcolo della mediana.

13. Lo scenario economico selezionato viene applicato direttamente al
    simulatore prima della valutazione delle soluzioni candidate, in modo
    che simulazione e ottimizzazione utilizzino la stessa struttura di costo.

14. I risultati possono essere salvati in Excel insieme alle impostazioni
    dell'ottimizzazione, allo scenario economico selezionato, ai costi
    effettivamente applicati, ai vincoli terminali degli accumuli e al
    riferimento diesel-only.


  !!!!!!!!  
Before the first run in Spyder, install pymoo in the same environment:
   pip install pymoo
"""

from pathlib import Path
import importlib.util
import tkinter as tk
from tkinter import filedialog, messagebox

import numpy as np
import pandas as pd

try:
    import matplotlib.pyplot as plt
    MATPLOTLIB_AVAILABLE = True
except ModuleNotFoundError:
    MATPLOTLIB_AVAILABLE = False

try:
    from pymoo.algorithms.moo.nsga2 import NSGA2
    from pymoo.algorithms.soo.nonconvex.ga import GA
    from pymoo.core.problem import ElementwiseProblem
    from pymoo.core.callback import Callback
    from pymoo.operators.crossover.sbx import SBX
    from pymoo.operators.mutation.pm import PM
    from pymoo.optimize import minimize
except ModuleNotFoundError as error:
    raise ModuleNotFoundError(
        "pymoo is not installed. In the Spyder console, run:\n"
        "pip install pymoo"
    ) from error



# USER SETTINGS


# Leave None to select the hourly rule-based simulator from a GUI window.
RULE_BASED_FILE = None

# First objective to minimise: "LCOE" or "NPC".
ECONOMIC_OBJECTIVE = "LCOE"

# Second objective:
#   "DECARBONISATION_PERCENT" -> MAXIMISE CO2 reduction vs diesel-only
#                                 (internally pymoo minimises its negative value)
#   "CO2"                    -> minimise annual CO2 emissions directly
#   "DIESEL_LOAD_PERCENT"    -> minimise annual load supplied by diesel
DECARBONISATION_OBJECTIVE = "DECARBONISATION_PERCENT"

# Reliability requirement for every candidate [%].
MAX_LPSP_PERCENT = 0.01

# Optional additional diesel contribution limit in hybrid modes [%].
# None = no extra limit; NSGA-II is free to explore the complete trade-off.
MAX_DIESEL_LOAD_PERCENT = None

# Multi-objective NSGA-II settings.
POPULATION_SIZE = 40
N_GENERATIONS = 60

# Epsilon-constraint settings. Each target is a separate single-objective
# pymoo GA run. The GUI can override the target range at execution time.
EPSILON_POPULATION_SIZE = 30
EPSILON_N_GENERATIONS = 40
EPSILON_MIN_PERCENT_DEFAULT = 0.0
EPSILON_MAX_PERCENT_DEFAULT = 100.0
EPSILON_STEP_PERCENT_DEFAULT = 10.0
EPSILON_TARGET_TOLERANCE_PERCENT = 1e-6

RANDOM_SEED = 42
CROSSOVER_PROBABILITY = 0.90
MUTATION_PROBABILITY = 0.15

# -----------------------------------------------------------------------------
# Economic scenario selector.
# Full numeric definitions live in the paired unified simulator so that direct
# simulation and optimisation always use exactly the same cost model.
# -----------------------------------------------------------------------------
ECONOMIC_SCENARIO_OPTIONS = {
    "LOW": "LOW - lower Arctic cost assumptions",
    "BASE": "BASE - central Arctic cost assumptions",
    "HIGH": "HIGH - upper Arctic cost assumptions",
    "MAROCCO_2022": "Marocco et al. 2022",
    "MCKINLEY_2025": "McKinley et al. 2025 (Alaska)",
    "JANKE_DEFAULT_2026": "Janke et al. 2026 (default)",
    "JANKE_SANIRAJAK_2026": "Janke et al. 2026 (Sanirajak)",
    "MEDIAN": "MEDIAN - median of the 4 literature presets",
}

# Selected in the initial GUI.
ECONOMIC_SCENARIO = None

# -----------------------------------------------------------------------------
# Search bounds used when Hornsund project constraints are OFF.
# They are numerical search-domain limits, not physical site limits.
# Increase a bound if the Pareto solutions repeatedly hit it.
# -----------------------------------------------------------------------------
FREE_PV_MAX_KW = 500.0
FREE_N_WT_MAX = 600
FREE_BATTERY_MAX_KWH = 3000.0
FREE_EL_MAX_KW = 500.0
FREE_H2_STORAGE_MAX_KWH = 50000.0
FREE_FC_MAX_KW = 500.0

# Battery interpretation for the Hornsund-constrained case.
# Project-scenario assumption adopted here:
#   constraints ON  -> planned Hornsund battery fixed at 500 kWh;
#   constraints OFF -> battery remains a free sizing variable.
FIX_HORNSUND_BATTERY_AT_PLANNED_SIZE = True

# Numerical H2 search bounds used with Hornsund constraints ON. Feasibility is
# still controlled by the H2-system budget from the simulator.
HORNSUND_EL_SEARCH_MAX_KW = 500.0
HORNSUND_H2_STORAGE_SEARCH_MAX_KWH = 20000.0
HORNSUND_FC_SEARCH_MAX_KW = 500.0

# Optional known designs. Leave [] for a fully random initial population.
# The order is always:
# PV, wind, battery, electrolyzer, H2 storage, fuel cell, diesel.
# Values are clipped to the active bounds, so fixed variables remain fixed.
INITIAL_DESIGNS = []

PLOT_PARETO_FRONT = True
PLOT_PARETO_DESIGN_SIZES = True

# These globals are configured only after the selected simulator is imported.
APPLY_HORNSUND_CONSTRAINTS = None
SYSTEM_MODE = None
OPTIMIZATION_METHOD = None       # "pareto" or "epsilon"
EPSILON_MIN_PERCENT = None
EPSILON_MAX_PERCENT = None
EPSILON_STEP_PERCENT = None
WT_RATED_POWER = None
N_WT_MAX_ACTIVE = None
SIZE_BOUNDS = None
ALL_VARIABLE_NAMES = None
VARIABLE_NAMES = None  # only variables with a non-zero search range
FIXED_SIZES = None
LOWER_BOUNDS = None
UPPER_BOUNDS = None
DIESEL_ONLY_REFERENCE = None


# =============================================================================
# GUI
# =============================================================================

def select_optimization_options():
    """Ask for constraints, architecture, economic scenario and optimisation method.

    This version uses a wider non-scrollable window, with the content split into
    two columns: general/system/economic options on the left, and optimisation
    method + epsilon settings on the right.
    """
    root = tk.Tk()
    root.title("Hornsund pymoo optimisation setup")
    root.resizable(True, True)
    root.attributes("-topmost", True)

    # Wider fixed-style layout: prefer horizontal development over scrolling.
    screen_w = root.winfo_screenwidth()
    screen_h = root.winfo_screenheight()
    window_w = min(1180, max(980, screen_w - 120))
    window_h = min(760, max(660, screen_h - 160))
    pos_x = max(0, (screen_w - window_w) // 2)
    pos_y = max(0, (screen_h - window_h) // 2)
    root.geometry(f"{window_w}x{window_h}+{pos_x}+{pos_y}")
    root.minsize(960, 640)

    apply_constraints_var = tk.BooleanVar(value=True)
    mode_var = tk.StringVar(value="mix")
    economic_scenario_var = tk.StringVar(value="BASE")
    method_var = tk.StringVar(value="epsilon")
    eps_min_var = tk.StringVar(value=str(EPSILON_MIN_PERCENT_DEFAULT))
    eps_max_var = tk.StringVar(value=str(EPSILON_MAX_PERCENT_DEFAULT))
    eps_step_var = tk.StringVar(value=str(EPSILON_STEP_PERCENT_DEFAULT))
    accepted = {"value": False}

    # Fixed bottom bar: Start/Cancel are always visible.
    button_frame = tk.Frame(root, padx=18, pady=10, bd=1, relief="groove")
    button_frame.pack(side="bottom", fill="x")

    # Main content area (non-scrollable).
    main_frame = tk.Frame(root, padx=18, pady=16)
    main_frame.pack(side="top", fill="both", expand=True)
    main_frame.grid_columnconfigure(0, weight=1, uniform="cols")
    main_frame.grid_columnconfigure(1, weight=1, uniform="cols")

    # -------------------------
    # Left column: constraints + system + economic scenario
    # -------------------------
    left_col = tk.Frame(main_frame)
    left_col.grid(row=0, column=0, sticky="nsew", padx=(0, 12))

    tk.Label(
        left_col,
        text="pymoo optimisation options",
        font=("Arial", 12, "bold"),
    ).pack(anchor="w", pady=(0, 10))

    tk.Checkbutton(
        left_col,
        text="Apply Hornsund project/design constraints",
        variable=apply_constraints_var,
    ).pack(anchor="w", pady=(0, 12))

    system_frame = tk.LabelFrame(
        left_col,
        text="System configuration",
        padx=10,
        pady=8,
    )
    system_frame.pack(fill="x")

    tk.Label(
        system_frame,
        text="Choose the system configuration to optimise:",
        font=("Arial", 10, "bold"),
    ).pack(anchor="w")

    mode_options = [
        ("Hybrid mix — Battery + H2 + diesel", "mix"),
        ("100% renewable — Battery + H2 (no diesel)", "renewable_only"),
        ("Hybrid mix — Battery only + diesel", "battery_only_mix"),
        ("100% renewable — Battery only (no diesel)", "battery_only_renewable"),
    ]
    for label, value in mode_options:
        tk.Radiobutton(
            system_frame,
            text=label,
            variable=mode_var,
            value=value,
        ).pack(anchor="w", padx=(12, 0), pady=(3, 0))

    economic_frame = tk.LabelFrame(
        left_col,
        text="Economic cost scenario",
        padx=10,
        pady=8,
    )
    economic_frame.pack(fill="x", pady=(14, 0))

    tk.Label(
        economic_frame,
        text="Choose the economic cost scenario used by the optimiser:",
        font=("Arial", 10, "bold"),
    ).pack(anchor="w")

    economic_options = [
        (label, key)
        for key, label in ECONOMIC_SCENARIO_OPTIONS.items()
    ]
    for label, value in economic_options:
        tk.Radiobutton(
            economic_frame,
            text=label,
            variable=economic_scenario_var,
            value=value,
        ).pack(anchor="w", padx=(12, 0), pady=(3, 0))

    tk.Label(
        economic_frame,
        text=(
            "LOW/BASE/HIGH use the updated Arctic sensitivity table. "
            "The four named literature cases use the complete harmonised cost "
            "structures defined in the paired simulator. MEDIAN is calculated "
            "parameter-by-parameter from the four literature presets only."
        ),
        fg="#555555",
        wraplength=500,
        justify="left",
    ).pack(anchor="w", pady=(6, 0))

    tk.Label(
        left_col,
        text=(
            "Battery rule: with Hornsund constraints ON the planned battery is "
            "fixed at 500 kWh; with constraints OFF its size is optimised."
        ),
        fg="#555555",
        wraplength=500,
        justify="left",
    ).pack(anchor="w", pady=(12, 0))

    # -------------------------
    # Right column: method + epsilon sweep
    # -------------------------
    right_col = tk.Frame(main_frame)
    right_col.grid(row=0, column=1, sticky="nsew", padx=(12, 0))

    method_frame = tk.LabelFrame(
        right_col,
        text="Optimisation method",
        padx=10,
        pady=8,
    )
    method_frame.pack(fill="x")

    tk.Label(
        method_frame,
        text="Choose the optimisation method:",
        font=("Arial", 10, "bold"),
    ).pack(anchor="w")

    tk.Radiobutton(
        method_frame,
        text="Multi-objective Pareto — NSGA-II",
        variable=method_var,
        value="pareto",
    ).pack(anchor="w", padx=(12, 0), pady=(3, 0))

    tk.Radiobutton(
        method_frame,
        text="Decarbonisation targets — epsilon-constraint",
        variable=method_var,
        value="epsilon",
    ).pack(anchor="w", padx=(12, 0), pady=(3, 0))

    epsilon_frame = tk.LabelFrame(
        right_col,
        text="Epsilon-constraint decarbonisation sweep",
        padx=10,
        pady=8,
    )
    epsilon_frame.pack(fill="x", pady=(14, 0))
    for col in range(6):
        epsilon_frame.grid_columnconfigure(col, weight=0)

    tk.Label(epsilon_frame, text="Minimum target [%]:").grid(
        row=0, column=0, sticky="e", padx=(0, 6), pady=2
    )
    eps_min_entry = tk.Entry(epsilon_frame, textvariable=eps_min_var, width=8)
    eps_min_entry.grid(row=0, column=1, sticky="w", pady=2)

    tk.Label(epsilon_frame, text="Maximum target [%]:").grid(
        row=0, column=2, sticky="e", padx=(16, 6), pady=2
    )
    eps_max_entry = tk.Entry(epsilon_frame, textvariable=eps_max_var, width=8)
    eps_max_entry.grid(row=0, column=3, sticky="w", pady=2)

    tk.Label(epsilon_frame, text="Step [%]:").grid(
        row=0, column=4, sticky="e", padx=(16, 6), pady=2
    )
    eps_step_entry = tk.Entry(epsilon_frame, textvariable=eps_step_var, width=8)
    eps_step_entry.grid(row=0, column=5, sticky="w", pady=2)

    tk.Label(
        epsilon_frame,
        text=(
            "Example: 0–100% with step 10% runs independent constrained "
            "optimisations at 0, 10, 20, …, 100% minimum decarbonisation."
        ),
        fg="#555555",
        wraplength=500,
        justify="left",
    ).grid(row=1, column=0, columnspan=6, sticky="w", pady=(6, 0))

    def update_epsilon_state(*_):
        state = "normal" if method_var.get() == "epsilon" else "disabled"
        for widget in (eps_min_entry, eps_max_entry, eps_step_entry):
            widget.configure(state=state)

    method_var.trace_add("write", update_epsilon_state)
    update_epsilon_state()

    def accept():
        if method_var.get() == "epsilon":
            try:
                eps_min = float(eps_min_var.get())
                eps_max = float(eps_max_var.get())
                eps_step = float(eps_step_var.get())
            except ValueError:
                messagebox.showerror(
                    "Invalid epsilon settings",
                    "Minimum, maximum and step must be numeric values.",
                    parent=root,
                )
                return

            if not (0.0 <= eps_min <= 100.0):
                messagebox.showerror(
                    "Invalid minimum target",
                    "Minimum decarbonisation must be between 0 and 100%.",
                    parent=root,
                )
                return
            if not (0.0 <= eps_max <= 100.0):
                messagebox.showerror(
                    "Invalid maximum target",
                    "Maximum decarbonisation must be between 0 and 100%.",
                    parent=root,
                )
                return
            if eps_max < eps_min:
                messagebox.showerror(
                    "Invalid target range",
                    "Maximum decarbonisation must be >= minimum decarbonisation.",
                    parent=root,
                )
                return
            if eps_step <= 0:
                messagebox.showerror(
                    "Invalid step",
                    "Decarbonisation step must be > 0.",
                    parent=root,
                )
                return
        else:
            eps_min = EPSILON_MIN_PERCENT_DEFAULT
            eps_max = EPSILON_MAX_PERCENT_DEFAULT
            eps_step = EPSILON_STEP_PERCENT_DEFAULT

        accepted["value"] = True
        accepted["apply_constraints"] = bool(apply_constraints_var.get())
        accepted["system_mode"] = mode_var.get()
        accepted["economic_scenario"] = economic_scenario_var.get()
        accepted["method"] = method_var.get()
        accepted["eps_min"] = eps_min
        accepted["eps_max"] = eps_max
        accepted["eps_step"] = eps_step
        root.destroy()

    def cancel():
        root.destroy()

    tk.Button(button_frame, text="Start", width=12, command=accept).pack(
        side="right", padx=(6, 0)
    )
    tk.Button(button_frame, text="Cancel", width=12, command=cancel).pack(
        side="right"
    )

    root.bind("<Return>", lambda _event: accept())
    root.bind("<Escape>", lambda _event: cancel())
    root.protocol("WM_DELETE_WINDOW", cancel)
    root.mainloop()

    if not accepted["value"]:
        raise SystemExit("Optimisation cancelled by user.")

    return (
        bool(accepted["apply_constraints"]),
        accepted["system_mode"],
        accepted["economic_scenario"],
        accepted["method"],
        float(accepted["eps_min"]),
        float(accepted["eps_max"]),
        float(accepted["eps_step"]),
    )

def select_rule_based_file():
    """Open a GUI and return the selected Python rule-based simulator."""
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        selected_file = filedialog.askopenfilename(
            parent=root,
            title="Select the rule-based Hornsund simulator",
            filetypes=[
                ("Python files", "*.py"),
                ("All files", "*.*"),
            ],
        )
    finally:
        root.destroy()

    if not selected_file:
        raise SystemExit("No simulator selected. Optimisation cancelled.")

    return Path(selected_file)


# =============================================================================
# MODEL IMPORT AND SEARCH SPACE
# =============================================================================

def import_rule_based_model(file_path):
    """Load the simulator without executing its final ``__main__`` section."""
    specification = importlib.util.spec_from_file_location(
        "rule_based_model", file_path
    )
    model = importlib.util.module_from_spec(specification)
    specification.loader.exec_module(model)
    return model


def apply_selected_economic_scenario(model, scenario_key):
    """Apply one of the eight scenarios through the paired simulator.

    The simulator owns the complete economic definitions (including size scaling,
    variable O&M and H2 auxiliary costs). Keeping those definitions in one place
    prevents direct runs and optimisation runs from drifting apart.
    """
    key = str(scenario_key).upper().strip()

    if key not in ECONOMIC_SCENARIO_OPTIONS:
        raise ValueError(
            f"Unknown economic scenario {scenario_key!r}. Choose one of: "
            + ", ".join(ECONOMIC_SCENARIO_OPTIONS)
        )

    if not hasattr(model, "COST_SCENARIOS") or not hasattr(model, "apply_cost_scenario"):
        raise AttributeError(
            "The selected simulator does not expose the unified COST_SCENARIOS / "
            "apply_cost_scenario interface. Select "
            "simulazione_offgrid_Hornsund_8_COST_SCENARIOS_STORAGE_END.py."
        )

    if key not in model.COST_SCENARIOS:
        raise ValueError(
            f"Scenario {key!r} is not available in the selected simulator. "
            f"Available: {', '.join(model.COST_SCENARIOS)}"
        )

    model.apply_cost_scenario(key)

    # Metadata retained in optimiser outputs.
    model.ECONOMIC_SCENARIO = key

    print("\nECONOMIC SCENARIO APPLIED BY OPTIMISER")
    print(f"  Scenario key:   {key}")
    print(f"  Scenario label: {getattr(model, 'COST_SCENARIO', key)}")
    print(f"  EL size scaling: {getattr(model, 'EL_USE_SIZE_SCALING', False)}")
    print(f"  FC size scaling: {getattr(model, 'FC_USE_SIZE_SCALING', False)}")

    fields = [
        "PV_CAPEX", "WIND_CAPEX", "BATTERY_CAPEX", "EL_CAPEX",
        "H2_TANK_CAPEX", "FC_CAPEX", "DIESEL_FUEL_COST",
    ]
    applied = {}
    for field in fields:
        if hasattr(model, field):
            applied[field] = float(getattr(model, field))
            print(f"  {field:34s} = {applied[field]:g}")
    return applied

def configure_search_space(model, apply_constraints, system_mode):
    """Build all sizing bounds consistently with the selected simulator."""
    global WT_RATED_POWER, N_WT_MAX_ACTIVE, SIZE_BOUNDS
    global ALL_VARIABLE_NAMES, VARIABLE_NAMES, FIXED_SIZES
    global LOWER_BOUNDS, UPPER_BOUNDS

    WT_RATED_POWER = float(model.WT_RATED_POWER)
    diesel_unit_size = float(model.DIESEL_UNIT_SIZE)

    if WT_RATED_POWER <= 0:
        raise ValueError("The simulator WT_RATED_POWER must be > 0 kW/turbine.")

    if apply_constraints:
        # Hornsund site/project assumptions are read from the simulator.
        pv_max = float(model.PV_MAX_SIZE)
        n_wt_max = int(model.N_WT_MAX)

        # Hornsund constrained scenario: the planned battery is treated as a
        # fixed project choice. Prefer the dedicated simulator value when present.
        battery_project_size = float(
            getattr(model, "BATTERY_HORNSUND_FIXED_SIZE", model.BATTERY_MAX_SIZE)
        )
        if FIX_HORNSUND_BATTERY_AT_PLANNED_SIZE:
            battery_bounds = (battery_project_size, battery_project_size)
        else:
            battery_bounds = (0.0, float(model.BATTERY_MAX_SIZE))

        el_max = HORNSUND_EL_SEARCH_MAX_KW
        h2_storage_max = HORNSUND_H2_STORAGE_SEARCH_MAX_KWH
        fc_max = HORNSUND_FC_SEARCH_MAX_KW
    else:
        # Wider numerical exploration, independent of site/project limits.
        pv_max = FREE_PV_MAX_KW
        n_wt_max = FREE_N_WT_MAX
        battery_bounds = (0.0, FREE_BATTERY_MAX_KWH)
        el_max = FREE_EL_MAX_KW
        h2_storage_max = FREE_H2_STORAGE_MAX_KWH
        fc_max = FREE_FC_MAX_KW

    N_WT_MAX_ACTIVE = n_wt_max

    diesel_modes = {"mix", "battery_only_mix"}
    renewable_only_modes = {"renewable_only", "battery_only_renewable"}
    battery_only_modes = {"battery_only_mix", "battery_only_renewable"}

    if system_mode in renewable_only_modes:
        diesel_bounds = (0.0, 0.0)
    elif system_mode in diesel_modes:
        # The installed genset is not a new sizing variable.
        diesel_bounds = (diesel_unit_size, diesel_unit_size)
    else:
        raise ValueError(f"Unknown SYSTEM_MODE: {system_mode}")

    if system_mode in battery_only_modes:
        # Exclude the complete hydrogen chain from the design space.
        el_bounds = (0.0, 0.0)
        h2_storage_bounds = (0.0, 0.0)
        fc_bounds = (0.0, 0.0)
    else:
        el_bounds = (0.0, el_max)
        h2_storage_bounds = (0.0, h2_storage_max)
        fc_bounds = (0.0, fc_max)

    SIZE_BOUNDS = {
        "pv": (0.0, pv_max),
        "wind": (0.0, n_wt_max * WT_RATED_POWER),
        "battery": battery_bounds,
        "electrolyzer": el_bounds,
        "hydrogen_storage": h2_storage_bounds,
        "fuel_cell": fc_bounds,
        "diesel": diesel_bounds,
    }

    ALL_VARIABLE_NAMES = tuple(SIZE_BOUNDS)

    # Do not pass fixed dimensions (lower == upper) to SBX/PM. Polynomial
    # mutation divides by the variable range, so zero-range variables can lead
    # to numerical problems. Fixed project choices are reconstructed in every
    # candidate instead.
    FIXED_SIZES = {
        name: float(lower)
        for name, (lower, upper) in SIZE_BOUNDS.items()
        if abs(upper - lower) <= 1e-12
    }
    VARIABLE_NAMES = tuple(
        name
        for name, (lower, upper) in SIZE_BOUNDS.items()
        if abs(upper - lower) > 1e-12
    )
    LOWER_BOUNDS = np.array(
        [SIZE_BOUNDS[name][0] for name in VARIABLE_NAMES], dtype=float
    )
    UPPER_BOUNDS = np.array(
        [SIZE_BOUNDS[name][1] for name in VARIABLE_NAMES], dtype=float
    )

    # Keep simulator metadata aligned with the outer optimisation mode.
    model.SYSTEM_MODE = system_mode
    model.APPLY_HORNSUND_CONSTRAINTS = bool(apply_constraints)

    mode_labels = {
        "mix": "HYBRID MIX — BATTERY + H2 + DIESEL",
        "renewable_only": "100% RENEWABLE — BATTERY + H2 / NO DIESEL",
        "battery_only_mix": "HYBRID MIX — BATTERY ONLY + DIESEL",
        "battery_only_renewable": "100% RENEWABLE — BATTERY ONLY / NO DIESEL",
    }

    print("\n" + "=" * 78)
    print("MULTI-OBJECTIVE OPTIMISATION SETUP")
    print("=" * 78)
    print(
        "Hornsund constraints: "
        + ("ENABLED" if apply_constraints else "DISABLED")
    )
    print("System mode:          " + mode_labels[system_mode])
    print(f"Economic scenario:    {ECONOMIC_SCENARIO}")
    print(f"Optimisation method:  {OPTIMIZATION_METHOD}")
    print(f"Economic objective:   {ECONOMIC_OBJECTIVE}")
    if OPTIMIZATION_METHOD == "pareto":
        print(f"Decarb objective:     {DECARBONISATION_OBJECTIVE}")
    else:
        print(
            f"Epsilon targets:      {EPSILON_MIN_PERCENT:g} to "
            f"{EPSILON_MAX_PERCENT:g}% by {EPSILON_STEP_PERCENT:g}%"
        )
    print(f"Max LPSP:             {MAX_LPSP_PERCENT:.4f} %")
    print("Storage end states:   ACTIVE (SOC_bat,end >= initial; LOH_H2,end >= initial)")
    print(f"WT rated power:       {WT_RATED_POWER:.3f} kW/turbine")
    print(f"Max turbines active:  {N_WT_MAX_ACTIVE:d}")

    print("\nSEARCH BOUNDS")
    for name, (lower, upper) in SIZE_BOUNDS.items():
        if name == "wind":
            unit = "kW"
        elif name in {"battery", "hydrogen_storage"}:
            unit = "kWh"
        else:
            unit = "kW"
        fixed = " [FIXED]" if abs(upper - lower) <= 1e-12 else ""
        print(
            f"  {name.replace('_', ' ').title():24s} "
            f"{lower:10.1f} - {upper:10.1f} {unit}{fixed}"
        )

    print(
        f"\nThe optimiser evolves {len(VARIABLE_NAMES)} variables: "
        + (", ".join(VARIABLE_NAMES) if VARIABLE_NAMES else "none")
    )
    if FIXED_SIZES:
        print(
            "Fixed sizes reconstructed outside the genetic operators: "
            + ", ".join(f"{k}={v:g}" for k, v in FIXED_SIZES.items())
        )

    if apply_constraints:
        if FIX_HORNSUND_BATTERY_AT_PLANNED_SIZE:
            print(
                "\nNOTE: the Hornsund battery is treated as a committed fixed "
                f"choice at {SIZE_BOUNDS['battery'][0]:.1f} kWh."
            )
        else:
            print(
                "\nNOTE: the Hornsund battery is NOT assumed already installed; "
                f"it is optimised from 0 to {SIZE_BOUNDS['battery'][1]:.1f} kWh."
            )
    else:
        print(
            "\nNOTE: constraints are OFF. The upper bounds are numerical search "
            "limits, not Hornsund site limits."
        )

    if system_mode in renewable_only_modes:
        print(
            "NOTE: in a 100% renewable architecture diesel CO2 is always zero. "
            "Therefore a cost-vs-decarbonisation Pareto front is degenerate; "
            "the useful comparison is mainly economic/design feasibility."
        )


# =============================================================================
# SIZING HELPERS
# =============================================================================

def sizes_from_vector(vector):
    """Reconstruct all seven sizes and quantise wind to whole turbines."""
    vector = np.asarray(vector, dtype=float)
    if len(VARIABLE_NAMES) > 0:
        vector = np.clip(vector, LOWER_BOUNDS, UPPER_BOUNDS)

    # Start from fixed project choices, then fill the dimensions actually
    # evolved by NSGA-II.
    sizes = {name: float(value) for name, value in FIXED_SIZES.items()}
    sizes.update(dict(zip(VARIABLE_NAMES, np.round(vector, 1))))

    # Every expected component must now be present.
    missing = [name for name in ALL_VARIABLE_NAMES if name not in sizes]
    if missing:
        raise RuntimeError(f"Missing reconstructed sizing variables: {missing}")

    n_wt = int(round(float(sizes["wind"]) / WT_RATED_POWER))
    n_wt_min = int(np.ceil(SIZE_BOUNDS["wind"][0] / WT_RATED_POWER - 1e-12))
    n_wt_max = int(np.floor(SIZE_BOUNDS["wind"][1] / WT_RATED_POWER + 1e-12))
    n_wt = int(np.clip(n_wt, n_wt_min, n_wt_max))
    sizes["wind"] = n_wt * WT_RATED_POWER

    return sizes


def apply_sizes(model, sizes):
    """Apply one candidate to the simulator, keeping all coupled globals aligned."""
    model.PV_SIZE = float(sizes["pv"])

    model.N_WT = int(round(float(sizes["wind"]) / float(model.WT_RATED_POWER)))
    model.WIND_SIZE = model.N_WT * float(model.WT_RATED_POWER)

    model.BATTERY_SIZE = float(sizes["battery"])

    if SYSTEM_MODE in {"battery_only_mix", "battery_only_renewable"}:
        model.EL_SIZE = 0.0
        model.H2_STORAGE_SIZE = 0.0
        model.FC_SIZE = 0.0
    else:
        model.EL_SIZE = float(sizes["electrolyzer"])
        model.H2_STORAGE_SIZE = float(sizes["hydrogen_storage"])
        model.FC_SIZE = float(sizes["fuel_cell"])

    model.DIESEL_SIZE = float(sizes["diesel"])
    model.SYSTEM_MODE = SYSTEM_MODE
    model.APPLY_HORNSUND_CONSTRAINTS = bool(APPLY_HORNSUND_CONSTRAINTS)


def calculate_diesel_only_reference(model, load, wind_cf, pv_cf):
    """Simulate the installed diesel-only reference once and restore the model."""
    saved = {
        "PV_SIZE": model.PV_SIZE,
        "N_WT": model.N_WT,
        "WIND_SIZE": model.WIND_SIZE,
        "BATTERY_SIZE": model.BATTERY_SIZE,
        "EL_SIZE": model.EL_SIZE,
        "H2_STORAGE_SIZE": model.H2_STORAGE_SIZE,
        "FC_SIZE": model.FC_SIZE,
        "DIESEL_SIZE": model.DIESEL_SIZE,
        "SYSTEM_MODE": getattr(model, "SYSTEM_MODE", None),
        "APPLY_HORNSUND_CONSTRAINTS": getattr(
            model, "APPLY_HORNSUND_CONSTRAINTS", None
        ),
    }

    try:
        model.PV_SIZE = 0.0
        model.N_WT = 0
        model.WIND_SIZE = 0.0
        model.BATTERY_SIZE = 0.0
        model.EL_SIZE = 0.0
        model.H2_STORAGE_SIZE = 0.0
        model.FC_SIZE = 0.0
        model.DIESEL_SIZE = float(model.DIESEL_UNIT_SIZE)
        model.SYSTEM_MODE = "diesel_only"
        model.APPLY_HORNSUND_CONSTRAINTS = False

        results = model.simulate(load, wind_cf, pv_cf)
    finally:
        for key, value in saved.items():
            setattr(model, key, value)

    fuel_l = float(results["diesel_fuel"].sum())
    co2_t = fuel_l * float(model.DIESEL_EMISSION_FACTOR) / 1000.0
    total_load_kwh = float(results["load"].sum())
    unserved_kwh = float(results["unserved"].sum())
    lpsp_percent = (
        100.0 * unserved_kwh / total_load_kwh if total_load_kwh > 0 else 0.0
    )

    return {
        "fuel_l_per_year": fuel_l,
        "co2_tonnes_per_year": co2_t,
        "lpsp_percent": lpsp_percent,
        "unserved_kwh": unserved_kwh,
    }


def evaluate_sizes(model, load, wind_cf, pv_cf, sizes):
    """Run one annual candidate and calculate objectives and all constraints."""
    apply_sizes(model, sizes)

    with np.errstate(divide="ignore", invalid="ignore"):
        results = model.simulate(load, wind_cf, pv_cf)

    metrics = model.calculate_operational_metrics(results)
    economics = model.calculate_economics(results, metrics)

    total_load = float(results["load"].sum())
    diesel_percent = (
        100.0 * float(results["diesel_to_load"].sum()) / total_load
        if total_load > 0 else 0.0
    )
    lpsp_percent = (
        100.0 * float(results["unserved"].sum()) / total_load
        if total_load > 0 else 0.0
    )
    diesel_fuel_l = float(results["diesel_fuel"].sum())
    co2_tonnes = diesel_fuel_l * float(model.DIESEL_EMISSION_FACTOR) / 1000.0

    baseline_co2 = float(DIESEL_ONLY_REFERENCE["co2_tonnes_per_year"])
    if baseline_co2 > 0:
        decarbonisation_percent = 100.0 * (1.0 - co2_tonnes / baseline_co2)
    else:
        decarbonisation_percent = np.nan

    npc_meur = float(economics["summary"]["Total NPC [M€]"])
    lcoe_eur_kwh = float(economics["summary"]["LCOE [€/kWh]"])

    if ECONOMIC_OBJECTIVE.upper() == "LCOE":
        economic_value = lcoe_eur_kwh
    elif ECONOMIC_OBJECTIVE.upper() == "NPC":
        economic_value = npc_meur
    else:
        raise ValueError('ECONOMIC_OBJECTIVE must be "LCOE" or "NPC".')

    decarb_choice = DECARBONISATION_OBJECTIVE.upper()
    if decarb_choice == "DECARBONISATION_PERCENT":
        # pymoo minimises all objectives, so maximise decarbonisation by
        # minimising its negative value.
        objective_2 = -decarbonisation_percent
    elif decarb_choice == "CO2":
        objective_2 = co2_tonnes
    elif decarb_choice == "DIESEL_LOAD_PERCENT":
        objective_2 = diesel_percent
    else:
        raise ValueError(
            'DECARBONISATION_OBJECTIVE must be "DECARBONISATION_PERCENT", '
            '"CO2" or "DIESEL_LOAD_PERCENT".'
        )

    # -------------------------------------------------------------------------
    # Reliability constraint: always active.
    # -------------------------------------------------------------------------
    lpsp_violation = (
        (lpsp_percent - MAX_LPSP_PERCENT) / max(MAX_LPSP_PERCENT, 1e-12)
    )

    # -------------------------------------------------------------------------
    # Annual storage sustainability constraints.
    # pymoo uses G <= 0, therefore a positive value means that the final storage
    # state is below the initial one and the candidate is NOT FEASIBLE.
    # -------------------------------------------------------------------------
    battery_soc_tolerance = float(
        getattr(model, "BATTERY_SOC_END_TOLERANCE", 1e-6)
    )
    if float(sizes["battery"]) > 0:
        battery_soc_initial = float(
            results.attrs.get("battery_soc_initial", model.BATTERY_SOC_INITIAL)
        )
        battery_soc_final = float(results.attrs["battery_soc_final"])
        battery_soc_difference = battery_soc_final - battery_soc_initial
        battery_soc_violation = (
            battery_soc_initial - battery_soc_final - battery_soc_tolerance
        )
    else:
        battery_soc_initial = 0.0
        battery_soc_final = 0.0
        battery_soc_difference = 0.0
        battery_soc_violation = -1.0

    h2_loh_tolerance = float(
        getattr(model, "H2_LOH_END_TOLERANCE", 1e-6)
    )
    if float(sizes["hydrogen_storage"]) > 0:
        h2_loh_initial = float(
            results.attrs.get(
                "h2_loh_initial",
                results.attrs.get("h2_soc_initial", model.H2_SOC_INITIAL),
            )
        )
        h2_loh_final = float(
            results.attrs.get("h2_loh_final", results.attrs["h2_soc_final"])
        )
        h2_loh_difference = h2_loh_final - h2_loh_initial
        h2_loh_violation = (
            h2_loh_initial - h2_loh_final - h2_loh_tolerance
        )
    else:
        h2_loh_initial = 0.0
        h2_loh_final = 0.0
        h2_loh_difference = 0.0
        h2_loh_violation = -1.0

    # Optional diesel-share constraint, meaningful only when diesel is active.
    if MAX_DIESEL_LOAD_PERCENT is not None and SYSTEM_MODE in {
        "mix", "battery_only_mix"
    }:
        diesel_share_violation = (
            diesel_percent - float(MAX_DIESEL_LOAD_PERCENT)
        ) / max(float(MAX_DIESEL_LOAD_PERCENT), 1.0)
    else:
        diesel_share_violation = -1.0

    # calculate_design_constraints() is also useful for H2 CAPEX reporting.
    design_constraints = model.calculate_design_constraints()
    h2_system_capex_eur = float(design_constraints["h2_system_capex"])

    # Hornsund constraints are activated only when requested.
    if APPLY_HORNSUND_CONSTRAINTS:
        h2_budget_violation = (
            -float(design_constraints["h2_budget_margin"])
            / max(float(model.H2_SYSTEM_BUDGET), 1e-12)
        )
        fuel_violation = (
            diesel_fuel_l - float(model.MAX_DIESEL_FUEL_L_PER_YEAR)
        ) / max(float(model.MAX_DIESEL_FUEL_L_PER_YEAR), 1e-12)

        # PV, wind and battery are already bounded explicitly. This catches any
        # additional checkable project constraint, e.g. crane-item mass when
        # component masses become available in the simulator.
        if not bool(design_constraints["overall_pre_feasible"]):
            pv_ok = bool(design_constraints["pv_size_feasible"])
            wind_ok = bool(design_constraints["wind_size_feasible"])
            battery_ok = bool(design_constraints["battery_size_feasible"])
            h2_ok = bool(design_constraints["h2_budget_feasible"])
            other_hornsund_violation = (
                1.0 if pv_ok and wind_ok and battery_ok and h2_ok else -1.0
            )
        else:
            other_hornsund_violation = -1.0
    else:
        h2_budget_violation = -1.0
        fuel_violation = -1.0
        other_hornsund_violation = -1.0

    constraint_values = {
        "lpsp_violation": float(lpsp_violation),
        "battery_soc_violation": float(battery_soc_violation),
        "h2_loh_violation": float(h2_loh_violation),
        "diesel_share_violation": float(diesel_share_violation),
        "h2_budget_violation": float(h2_budget_violation),
        "fuel_violation": float(fuel_violation),
        "other_hornsund_violation": float(other_hornsund_violation),
    }

    feasible = (
        constraint_values["lpsp_violation"] <= 1e-12
        and constraint_values["battery_soc_violation"] <= 1e-12
        and constraint_values["h2_loh_violation"] <= 1e-12
        and (
            MAX_DIESEL_LOAD_PERCENT is None
            or SYSTEM_MODE not in {"mix", "battery_only_mix"}
            or constraint_values["diesel_share_violation"] <= 1e-12
        )
        and (
            not APPLY_HORNSUND_CONSTRAINTS
            or (
                constraint_values["h2_budget_violation"] <= 1e-12
                and constraint_values["fuel_violation"] <= 1e-12
                and constraint_values["other_hornsund_violation"] <= 1e-12
            )
        )
    )

    return {
        **sizes,
        "n_wt": int(model.N_WT),
        "economic_objective": economic_value,
        "pymoo_objective_2": objective_2,
        "npc_meur": npc_meur,
        "lcoe_eur_kwh": lcoe_eur_kwh,
        "co2_tonnes_per_year": co2_tonnes,
        "decarbonisation_percent": decarbonisation_percent,
        "diesel_percent": diesel_percent,
        "lpsp_percent": lpsp_percent,
        "unserved_kwh": float(results["unserved"].sum()),
        "diesel_fuel_l_per_year": diesel_fuel_l,
        "battery_soc_initial": battery_soc_initial,
        "battery_soc_final": battery_soc_final,
        "battery_soc_difference": battery_soc_difference,
        "battery_soc_tolerance": battery_soc_tolerance,
        "h2_loh_initial": h2_loh_initial,
        "h2_loh_final": h2_loh_final,
        "h2_loh_difference": h2_loh_difference,
        "h2_loh_tolerance": h2_loh_tolerance,
        "h2_system_capex_eur": h2_system_capex_eur,
        "hornsund_constraints_enabled": bool(APPLY_HORNSUND_CONSTRAINTS),
        "system_mode": SYSTEM_MODE,
        **constraint_values,
        "feasible": bool(feasible),
    }


# =============================================================================
# PYMOO PROBLEM
# =============================================================================

def active_constraint_names(decarbonisation_target_percent=None):
    """Return the exact inequality constraints passed to pymoo.

    Storage sustainability is always enforced when the corresponding storage
    component is active. Inactive storage receives a negative dummy violation.
    """
    names = [
        "lpsp_violation",
        "battery_soc_violation",
        "h2_loh_violation",
    ]

    if MAX_DIESEL_LOAD_PERCENT is not None and SYSTEM_MODE in {
        "mix", "battery_only_mix"
    }:
        names.append("diesel_share_violation")

    if APPLY_HORNSUND_CONSTRAINTS:
        names.extend([
            "h2_budget_violation",
            "fuel_violation",
            "other_hornsund_violation",
        ])

    if decarbonisation_target_percent is not None:
        names.append("decarbonisation_target_violation")

    return names


class OffGridSizingProblem(ElementwiseProblem):
    """Connection between pymoo and the annual Hornsund simulator."""

    def __init__(
        self,
        model,
        load,
        wind_cf,
        pv_cf,
        optimization_method="pareto",
        decarbonisation_target_percent=None,
        shared_cache=None,
    ):
        self.optimization_method = optimization_method
        self.decarbonisation_target_percent = decarbonisation_target_percent
        self.constraint_names = active_constraint_names(
            decarbonisation_target_percent
        )

        n_obj = 2 if optimization_method == "pareto" else 1

        super().__init__(
            n_var=len(VARIABLE_NAMES),
            n_obj=n_obj,
            n_ieq_constr=len(self.constraint_names),
            xl=LOWER_BOUNDS,
            xu=UPPER_BOUNDS,
        )
        self.model = model
        self.load = load
        self.wind_cf = wind_cf
        self.pv_cf = pv_cf
        self.cache = shared_cache if shared_cache is not None else {}

    def record_for_vector(self, vector):
        """Evaluate each rounded design once and cache repeated candidates."""
        sizes = sizes_from_vector(vector)
        key = tuple(sizes[name] for name in ALL_VARIABLE_NAMES)
        if key not in self.cache:
            self.cache[key] = evaluate_sizes(
                self.model, self.load, self.wind_cf, self.pv_cf, sizes
            )
        return self.cache[key]

    def _evaluate(self, vector, out, *args, **kwargs):
        record = self.record_for_vector(vector)

        if self.optimization_method == "pareto":
            out["F"] = [
                record["economic_objective"],
                record["pymoo_objective_2"],
            ]
        else:
            # Epsilon-constraint formulation:
            # minimise only the economic objective while decarbonisation is a
            # lower-bound inequality constraint.
            out["F"] = [record["economic_objective"]]

        constraint_values = {
            name: record[name]
            for name in self.constraint_names
            if name != "decarbonisation_target_violation"
        }

        if self.decarbonisation_target_percent is not None:
            achieved = float(record["decarbonisation_percent"])
            target = float(self.decarbonisation_target_percent)
            # pymoo requires G <= 0. Divide by 100 only for numerical scaling.
            constraint_values["decarbonisation_target_violation"] = (
                target - achieved
            ) / 100.0

        out["G"] = [
            constraint_values[name] for name in self.constraint_names
        ]

# =============================================================================
# PYMOO ALGORITHMS
# =============================================================================

def initial_population(population_size, seed, warm_start_vector=None):
    """Create a population and optionally insert known/warm-start designs."""
    if len(INITIAL_DESIGNS) > population_size:
        raise ValueError(
            "INITIAL_DESIGNS cannot contain more rows than the population size."
        )

    random = np.random.default_rng(seed)
    population = random.uniform(
        LOWER_BOUNDS,
        UPPER_BOUNDS,
        (population_size, len(VARIABLE_NAMES)),
    )

    next_row = 0
    if warm_start_vector is not None and population_size > 0:
        warm_start_vector = np.asarray(warm_start_vector, dtype=float)
        if len(warm_start_vector) == len(VARIABLE_NAMES):
            population[0] = np.clip(
                warm_start_vector, LOWER_BOUNDS, UPPER_BOUNDS
            )
            next_row = 1

    for design in INITIAL_DESIGNS:
        if next_row >= population_size:
            break
        design = np.asarray(design, dtype=float)
        if len(design) != len(ALL_VARIABLE_NAMES):
            raise ValueError(
                "Each INITIAL_DESIGNS row must contain seven size values in the "
                "documented full-size order."
            )
        full_design = dict(zip(ALL_VARIABLE_NAMES, design))
        active_design = np.array(
            [full_design[name] for name in VARIABLE_NAMES], dtype=float
        )
        population[next_row] = np.clip(
            active_design, LOWER_BOUNDS, UPPER_BOUNDS
        )
        next_row += 1

    return population


def _result_vectors(result, include_population=False):
    """Return candidate vectors robustly from a pymoo result."""
    vectors = []

    if result.X is not None:
        x = np.asarray(result.X, dtype=float)
        if x.ndim == 1:
            x = x.reshape(1, -1)
        vectors.append(x)

    if include_population and result.pop is not None:
        pop_x = np.asarray(result.pop.get("X"), dtype=float)
        if pop_x.ndim == 1:
            pop_x = pop_x.reshape(1, -1)
        vectors.append(pop_x)

    if not vectors:
        return np.empty((0, len(VARIABLE_NAMES)), dtype=float)

    combined = np.vstack(vectors)
    if len(combined) <= 1:
        return combined

    return np.unique(np.round(combined, 12), axis=0)


def build_epsilon_targets(minimum, maximum, step):
    """Build an inclusive sequence of decarbonisation targets [%]."""
    minimum = float(minimum)
    maximum = float(maximum)
    step = float(step)

    if not (0.0 <= minimum <= maximum <= 100.0):
        raise ValueError("Epsilon targets must satisfy 0 <= min <= max <= 100.")
    if step <= 0:
        raise ValueError("Epsilon target step must be > 0.")

    targets = []
    current = minimum
    while current <= maximum + 1e-10:
        targets.append(round(current, 10))
        current += step

    if abs(targets[-1] - maximum) > 1e-9:
        targets.append(maximum)

    return np.asarray(targets, dtype=float)



class EpsilonProgressCallback(Callback):
    """Readable convergence table for the single-objective epsilon runs.

    IMPORTANT:
    ``epsilon_target`` is the physical minimum decarbonisation constraint [%].
    It is NOT pymoo's multi-objective ``eps`` convergence metric shown by
    NSGA-II. The latter is meaningful for Pareto-front convergence and is kept
    in the native verbose output of Pareto mode.
    """

    def __init__(self, problem, target_percent):
        super().__init__()
        self.problem = problem
        self.target_percent = float(target_percent)
        self.previous_best = None
        self.header_printed = False

    def _print_header(self):
        if self.header_printed:
            return
        print(
            "\n"
            + "-" * 112
        )
        print(
            f"{'n_gen':>6s} | {'n_eval':>7s} | {'epsilon_target':>14s} | "
            f"{'cv_min':>10s} | {'best_LCOE':>11s} | "
            f"{'decarb_best':>12s} | {'delta_best':>11s} | {'indicator':>12s}"
        )
        print("-" * 112)
        self.header_printed = True

    def notify(self, algorithm):
        self._print_header()

        population = algorithm.pop
        if population is None or len(population) == 0:
            return

        cv = np.asarray(population.get("CV"), dtype=float).reshape(-1)
        objectives = np.asarray(population.get("F"), dtype=float).reshape(-1)
        vectors = np.asarray(population.get("X"), dtype=float)

        cv_min = float(np.nanmin(cv)) if cv.size else np.nan
        feasible_mask = (
            np.isfinite(objectives)
            & np.isfinite(cv)
            & (cv <= 1e-12)
        )

        best_value = np.nan
        achieved_decarb = np.nan
        delta_best = np.nan
        indicator = "NO FEASIBLE"

        if np.any(feasible_mask):
            feasible_indices = np.where(feasible_mask)[0]
            local_best = int(
                feasible_indices[
                    np.argmin(objectives[feasible_indices])
                ]
            )

            best_value = float(objectives[local_best])
            best_vector = vectors[local_best]
            record = self.problem.record_for_vector(best_vector)
            achieved_decarb = float(record["decarbonisation_percent"])

            if self.previous_best is None:
                indicator = "FIRST FEAS."
            else:
                delta_best = abs(best_value - self.previous_best) / max(
                    abs(self.previous_best), 1e-12
                )
                if best_value < self.previous_best - 1e-10:
                    indicator = "IMPROVING"
                else:
                    indicator = "STABLE"

            self.previous_best = best_value

        n_eval = getattr(algorithm.evaluator, "n_eval", 0)

        def fmt(value, width, precision):
            if np.isfinite(value):
                return f"{value:{width}.{precision}f}"
            return f"{'-':>{width}s}"

        print(
            f"{int(algorithm.n_gen):6d} | "
            f"{int(n_eval):7d} | "
            f"{self.target_percent:14.2f} | "
            f"{fmt(cv_min, 10, 4)} | "
            f"{fmt(best_value, 11, 6)} | "
            f"{fmt(achieved_decarb, 12, 3)} | "
            f"{fmt(delta_best, 11, 3)} | "
            f"{indicator:>12s}"
        )


def prepare_optimisation():
    """Load simulator/data, configure bounds and calculate diesel reference."""
    global DIESEL_ONLY_REFERENCE

    rule_based_file = (
        select_rule_based_file()
        if RULE_BASED_FILE is None
        else Path(RULE_BASED_FILE)
    )

    if not rule_based_file.is_file():
        raise FileNotFoundError(f"Rule-based file not found: {rule_based_file}")

    print(f"Selected simulator: {rule_based_file}")
    model = import_rule_based_model(rule_based_file)

    # Apply the economic case chosen in the optimiser GUI directly through the
    # paired simulator. This avoids a second cost window and guarantees that
    # every candidate uses one coherent one-of-eight economic scenario.
    apply_selected_economic_scenario(model, ECONOMIC_SCENARIO)

    configure_search_space(
        model,
        APPLY_HORNSUND_CONSTRAINTS,
        SYSTEM_MODE,
    )

    load, wind_cf, pv_cf = model.load_profiles(
        model.EXCEL_FILE, model.SHEET_NAME
    )

    DIESEL_ONLY_REFERENCE = calculate_diesel_only_reference(
        model, load, wind_cf, pv_cf
    )
    print("\nDIESEL-ONLY REFERENCE")
    print(
        f"  Fuel: {DIESEL_ONLY_REFERENCE['fuel_l_per_year']:.1f} L/year"
    )
    print(
        f"  CO2:  {DIESEL_ONLY_REFERENCE['co2_tonnes_per_year']:.2f} tCO2/year"
    )
    print(
        f"  LPSP: {DIESEL_ONLY_REFERENCE['lpsp_percent']:.5f} %"
    )

    return model, load, wind_cf, pv_cf


def optimize_pareto(model, load, wind_cf, pv_cf):
    """Run the original two-objective NSGA-II Pareto optimisation."""
    problem = OffGridSizingProblem(
        model,
        load,
        wind_cf,
        pv_cf,
        optimization_method="pareto",
    )

    print("\nACTIVE PYMOO CONSTRAINTS")
    for name in problem.constraint_names:
        print(f"  - {name}")

    algorithm = NSGA2(
        pop_size=POPULATION_SIZE,
        sampling=initial_population(
            POPULATION_SIZE, RANDOM_SEED
        ),
        crossover=SBX(prob=CROSSOVER_PROBABILITY, eta=15),
        mutation=PM(prob=MUTATION_PROBABILITY, eta=20),
        eliminate_duplicates=True,
    )

    print(
        "\nNOTE: in Pareto mode, the native pymoo NSGA-II columns 'eps' and "
        "'indicator' are convergence diagnostics of the Pareto front; they are "
        "not decarbonisation targets."
    )

    result = minimize(
        problem,
        algorithm,
        termination=("n_gen", N_GENERATIONS),
        seed=RANDOM_SEED,
        verbose=True,
    )

    vectors = _result_vectors(result, include_population=False)
    records = [problem.record_for_vector(vector).copy() for vector in vectors]

    if not records:
        return pd.DataFrame()

    candidates = pd.DataFrame(records).drop_duplicates(
        subset=list(ALL_VARIABLE_NAMES)
    )
    pareto = candidates[candidates["feasible"]].copy()

    if not pareto.empty:
        pareto = pareto.sort_values(
            ["economic_objective", "pymoo_objective_2"]
        ).reset_index(drop=True)

    return pareto


def optimize_epsilon_curve(model, load, wind_cf, pv_cf):
    """Run one single-objective constrained optimisation per epsilon target."""
    targets = build_epsilon_targets(
        EPSILON_MIN_PERCENT,
        EPSILON_MAX_PERCENT,
        EPSILON_STEP_PERCENT,
    )

    print("\nEPSILON-CONSTRAINT TARGETS")
    print("  " + ", ".join(f"{target:g}%" for target in targets))
    print(
        f"  Each target: population={EPSILON_POPULATION_SIZE}, "
        f"generations={EPSILON_N_GENERATIONS}"
    )
    print(
        "  Constraint used at each point: "
        "Decarbonisation >= requested target."
    )

    if SYSTEM_MODE in {"renewable_only", "battery_only_renewable"}:
        print(
            "\nWARNING: the selected architecture has no diesel. Its operating "
            "decarbonisation is therefore 100% by construction, so a full "
            "epsilon sweep is usually redundant."
        )

    shared_cache = {}
    rows = []
    warm_start_vector = None

    for index, target in enumerate(targets):
        print("\n" + "=" * 78)
        print(
            f"EPSILON TARGET {index + 1}/{len(targets)}: "
            f"decarbonisation >= {target:.2f}%"
        )
        print("=" * 78)

        problem = OffGridSizingProblem(
            model,
            load,
            wind_cf,
            pv_cf,
            optimization_method="epsilon",
            decarbonisation_target_percent=float(target),
            shared_cache=shared_cache,
        )

        algorithm = GA(
            pop_size=EPSILON_POPULATION_SIZE,
            sampling=initial_population(
                EPSILON_POPULATION_SIZE,
                RANDOM_SEED + index,
                warm_start_vector=warm_start_vector,
            ),
            crossover=SBX(prob=CROSSOVER_PROBABILITY, eta=15),
            mutation=PM(prob=MUTATION_PROBABILITY, eta=20),
            eliminate_duplicates=True,
        )

        progress_callback = EpsilonProgressCallback(
            problem,
            target_percent=float(target),
        )

        result = minimize(
            problem,
            algorithm,
            termination=("n_gen", EPSILON_N_GENERATIONS),
            seed=RANDOM_SEED + index,
            callback=progress_callback,
            # The custom callback replaces pymoo's default single-objective
            # table so that the physical epsilon target is visible explicitly.
            verbose=False,
        )

        # For a constrained single-objective run inspect both the reported
        # optimum and the final population, then explicitly select the cheapest
        # design satisfying all base constraints and the epsilon target.
        vectors = _result_vectors(result, include_population=True)
        records = [
            problem.record_for_vector(vector).copy()
            for vector in vectors
        ]

        feasible_records = []
        for record in records:
            achieved = float(record["decarbonisation_percent"])
            target_ok = (
                achieved + EPSILON_TARGET_TOLERANCE_PERCENT >= float(target)
            )
            if bool(record["feasible"]) and target_ok:
                feasible_records.append(record)

        if feasible_records:
            best = min(
                feasible_records,
                key=lambda item: float(item["economic_objective"]),
            ).copy()

            best["minimum_decarbonisation_target_percent"] = float(target)
            best["decarbonisation_target_margin_percent"] = (
                float(best["decarbonisation_percent"]) - float(target)
            )
            best["epsilon_target_feasible"] = True
            rows.append(best)

            warm_start_vector = np.array(
                [best[name] for name in VARIABLE_NAMES],
                dtype=float,
            )

            print(
                f"  BEST FEASIBLE: {ECONOMIC_OBJECTIVE}="
                f"{best['economic_objective']:.6f}, "
                f"achieved decarb={best['decarbonisation_percent']:.3f}%"
            )
        else:
            rows.append({
                "minimum_decarbonisation_target_percent": float(target),
                "decarbonisation_target_margin_percent": np.nan,
                "epsilon_target_feasible": False,
                "economic_objective": np.nan,
                "lcoe_eur_kwh": np.nan,
                "npc_meur": np.nan,
                "decarbonisation_percent": np.nan,
                "co2_tonnes_per_year": np.nan,
                "diesel_percent": np.nan,
                "lpsp_percent": np.nan,
                "feasible": False,
            })
            print("  NO FEASIBLE DESIGN FOUND FOR THIS TARGET.")

    epsilon_results = pd.DataFrame(rows)
    return epsilon_results


def optimize():
    """Dispatch to Pareto NSGA-II or epsilon-constraint optimisation."""
    model, load, wind_cf, pv_cf = prepare_optimisation()

    if OPTIMIZATION_METHOD == "pareto":
        return optimize_pareto(model, load, wind_cf, pv_cf), model

    if OPTIMIZATION_METHOD == "epsilon":
        return optimize_epsilon_curve(model, load, wind_cf, pv_cf), model

    raise ValueError(
        'OPTIMIZATION_METHOD must be "pareto" or "epsilon".'
    )

# =============================================================================
# OUTPUT
# =============================================================================

def _selected_decarbonisation_label():
    choice = DECARBONISATION_OBJECTIVE.upper()
    if choice == "DECARBONISATION_PERCENT":
        return "Decarbonisation vs diesel-only [%]"
    if choice == "CO2":
        return "Annual CO2 emissions [tCO2/year]"
    return "Diesel contribution to load [%]"


def _selected_decarbonisation_column():
    choice = DECARBONISATION_OBJECTIVE.upper()
    if choice == "DECARBONISATION_PERCENT":
        return "decarbonisation_percent"
    if choice == "CO2":
        return "co2_tonnes_per_year"
    return "diesel_percent"


def print_pareto_summary(pareto):
    """Print the least-cost and strongest-decarbonisation feasible designs."""
    print(f"\nFEASIBLE PARETO SOLUTIONS: {len(pareto)}")
    if pareto.empty:
        print(
            "No feasible Pareto solution was found. Check which constraint is "
            "binding, increase numerical search bounds, or increase NSGA-II "
            "population/generations."
        )
        return

    columns = [
        "pv", "wind", "n_wt", "battery", "electrolyzer",
        "hydrogen_storage", "fuel_cell", "diesel",
        "lcoe_eur_kwh", "npc_meur", "co2_tonnes_per_year",
        "decarbonisation_percent", "diesel_percent", "lpsp_percent",
        "battery_soc_initial", "battery_soc_final",
        "h2_loh_initial", "h2_loh_final", "h2_system_capex_eur",
    ]

    least_cost = pareto.sort_values("economic_objective").iloc[0]
    print("\nLOWEST ECONOMIC OBJECTIVE")
    print(least_cost[columns].to_string())

    choice = DECARBONISATION_OBJECTIVE.upper()
    if choice == "DECARBONISATION_PERCENT":
        strongest = pareto.sort_values(
            "decarbonisation_percent", ascending=False
        ).iloc[0]
        title = "HIGHEST DECARBONISATION"
    elif choice == "CO2":
        strongest = pareto.sort_values("co2_tonnes_per_year").iloc[0]
        title = "LOWEST CO2"
    else:
        strongest = pareto.sort_values("diesel_percent").iloc[0]
        title = "LOWEST DIESEL CONTRIBUTION"

    print(f"\n{title}")
    print(strongest[columns].to_string())


def print_epsilon_summary(results):
    """Print the epsilon-constraint curve in a compact table."""
    print("\n" + "=" * 78)
    print("EPSILON-CONSTRAINT RESULTS")
    print("=" * 78)

    if results.empty:
        print("No epsilon-constraint results are available.")
        return

    columns = [
        "minimum_decarbonisation_target_percent",
        "decarbonisation_percent",
        "lcoe_eur_kwh",
        "npc_meur",
        "pv",
        "n_wt",
        "battery",
        "electrolyzer",
        "hydrogen_storage",
        "fuel_cell",
        "diesel_fuel_l_per_year",
        "lpsp_percent",
        "epsilon_target_feasible",
    ]
    available = [column for column in columns if column in results.columns]
    print(results[available].to_string(index=False))


def ask_save_results(results_table, model=None):
    """Ask whether and where to save settings and optimisation results."""
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    selected_file = None
    try:
        label = (
            "Pareto results"
            if OPTIMIZATION_METHOD == "pareto"
            else "epsilon-constraint results"
        )
        save = messagebox.askyesno(
            "Save pymoo results",
            f"Do you want to save the {label} to Excel?",
            parent=root,
        )
        if not save:
            print("\nResults were not saved.")
            return

        scenario_tag = str(ECONOMIC_SCENARIO or "COST").upper()
        initial_name = (
            f"optimal_sizing_multiobjective_pymoo_results_{scenario_tag}.xlsx"
            if OPTIMIZATION_METHOD == "pareto"
            else f"optimal_sizing_epsilon_constraint_results_{scenario_tag}.xlsx"
        )

        selected_file = filedialog.asksaveasfilename(
            parent=root,
            title="Save pymoo optimisation results",
            defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
            initialfile=initial_name,
        )
    finally:
        root.destroy()

    if not selected_file:
        print("\nSave cancelled. Results were not saved.")
        return

    settings_rows = [
        ("Optimisation method", OPTIMIZATION_METHOD),
        ("Economic objective", ECONOMIC_OBJECTIVE),
        ("Decarbonisation objective (Pareto mode)", DECARBONISATION_OBJECTIVE),
        ("Hornsund constraints enabled", APPLY_HORNSUND_CONSTRAINTS),
        ("System mode", SYSTEM_MODE),
        ("Economic scenario selected by optimiser", ECONOMIC_SCENARIO_OPTIONS.get(ECONOMIC_SCENARIO, ECONOMIC_SCENARIO)),
        ("Maximum LPSP [%]", MAX_LPSP_PERCENT),
        ("Maximum diesel contribution [%]", MAX_DIESEL_LOAD_PERCENT),
        ("NSGA-II population size", POPULATION_SIZE),
        ("NSGA-II number of generations", N_GENERATIONS),
        ("Epsilon GA population size", EPSILON_POPULATION_SIZE),
        ("Epsilon GA number of generations", EPSILON_N_GENERATIONS),
        ("Epsilon minimum target [%]", EPSILON_MIN_PERCENT),
        ("Epsilon maximum target [%]", EPSILON_MAX_PERCENT),
        ("Epsilon step [%]", EPSILON_STEP_PERCENT),
        (
            "Epsilon progress indicator",
            "custom: FIRST FEAS. / IMPROVING / STABLE / NO FEASIBLE",
        ),
        (
            "NSGA-II eps meaning",
            "Pareto convergence diagnostic; not decarbonisation epsilon",
        ),
        ("Random seed", RANDOM_SEED),
        ("Wind turbine rated power [kW]", WT_RATED_POWER),
        ("Maximum active wind turbines", N_WT_MAX_ACTIVE),
        (
            "Battery fixed at planned Hornsund size",
            FIX_HORNSUND_BATTERY_AT_PLANNED_SIZE,
        ),
        (
            "Battery terminal SOC constraint",
            "ACTIVE: SOC_end >= SOC_initial",
        ),
        (
            "H2 terminal LOH constraint",
            "ACTIVE: LOH_end >= LOH_initial",
        ),
        (
            "Diesel-only reference fuel [L/year]",
            DIESEL_ONLY_REFERENCE["fuel_l_per_year"],
        ),
        (
            "Diesel-only reference CO2 [tCO2/year]",
            DIESEL_ONLY_REFERENCE["co2_tonnes_per_year"],
        ),
    ]

    # Save the economic assumptions actually used by the imported simulator so
    # every Pareto / epsilon run remains reproducible, including manual overrides.
    if model is not None:
        settings_rows.append(("Economic scenario label", getattr(model, "COST_SCENARIO", "n.a.")))
        settings_rows.append(("Economic scenario key", getattr(model, "COST_SCENARIO_KEY", "n.a.")))
        economic_fields = [
            ("PV CAPEX [EUR/kW]", "PV_CAPEX"),
            ("PV fixed O&M [EUR/(kW y)]", "PV_FIXED_OM"),
            ("PV variable O&M [EUR/kWh]", "PV_VARIABLE_OM"),
            ("Wind CAPEX [EUR/kW]", "WIND_CAPEX"),
            ("Wind fixed O&M [EUR/(kW y)]", "WIND_FIXED_OM"),
            ("Wind variable O&M [EUR/kWh]", "WIND_VARIABLE_OM"),
            ("Battery CAPEX [EUR/kWh]", "BATTERY_CAPEX"),
            ("Battery replacement [EUR/kWh]", "BATTERY_REPLACEMENT_COST"),
            ("Battery fixed O&M [EUR/(kWh y)]", "BATTERY_FIXED_OM"),
            ("Battery variable O&M [EUR/kWh throughput]", "BATTERY_VARIABLE_OM"),
            ("Electrolyzer reference CAPEX [EUR/kW]", "EL_CAPEX"),
            ("Electrolyzer stack replacement [% CAPEX]", "EL_STACK_REPLACEMENT_FRACTION"),
            ("Electrolyzer fractional O&M [%/y]", "EL_TOTAL_OM_FRACTION"),
            ("Electrolyzer fixed O&M [EUR/(kW y)]", "EL_FIXED_OM_PER_KW_YEAR"),
            ("H2 tank CAPEX [EUR/kg]", "H2_TANK_CAPEX"),
            ("H2 tank replacement [EUR/kg]", "H2_TANK_REPLACEMENT_COST"),
            ("H2 tank fixed O&M [% CAPEX/y]", "H2_TANK_FIXED_OM_FRACTION"),
            ("H2 tank fixed O&M [EUR/(kg y)]", "H2_TANK_FIXED_OM_PER_KG_YEAR"),
            ("H2 water treatment CAPEX [EUR/kW_EL]", "H2_WATER_TREATMENT_CAPEX_PER_KW_EL"),
            ("H2 compressor CAPEX [EUR/kW_EL]", "H2_COMPRESSOR_CAPEX_PER_KW_EL"),
            ("H2 system variable O&M [EUR/kWh_H2 throughput]", "H2_SYSTEM_VARIABLE_OM"),
            ("Fuel cell reference CAPEX [EUR/kW]", "FC_CAPEX"),
            ("Fuel cell stack replacement [% CAPEX]", "FC_STACK_REPLACEMENT_FRACTION"),
            ("Fuel cell fractional O&M [%/y]", "FC_TOTAL_OM_FRACTION"),
            ("Fuel cell fixed O&M [EUR/(kW y)]", "FC_FIXED_OM_PER_KW_YEAR"),
            ("Fuel cell O&M [EUR/h]", "FC_OM_PER_OPERATING_HOUR"),
            ("Diesel replacement [EUR/kW]", "DIESEL_REPLACEMENT_COST"),
            ("Diesel fuel [EUR/L]", "DIESEL_FUEL_COST"),
            ("Diesel O&M [EUR/h]", "DIESEL_OM_PER_OPERATING_HOUR"),
        ]
        for label, attribute in economic_fields:
            if hasattr(model, attribute):
                value = getattr(model, attribute)
                if attribute.endswith("_FRACTION"):
                    value = 100.0 * float(value)
                settings_rows.append((label, value))

    for name, bounds in SIZE_BOUNDS.items():
        settings_rows.append((f"{name} lower bound", bounds[0]))
        settings_rows.append((f"{name} upper bound", bounds[1]))

    sheet_name = (
        "Pareto solutions"
        if OPTIMIZATION_METHOD == "pareto"
        else "Epsilon curve"
    )

    with pd.ExcelWriter(selected_file, engine="openpyxl") as writer:
        pd.DataFrame(
            settings_rows, columns=["Setting", "Value"]
        ).to_excel(writer, sheet_name="Settings", index=False)
        results_table.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"\nCreated: {Path(selected_file).resolve()}")

def plot_pareto_front(pareto):
    """Plot economic performance against the selected decarbonisation metric."""
    if not PLOT_PARETO_FRONT or pareto.empty:
        return
    if not MATPLOTLIB_AVAILABLE:
        print("Plot not created: install matplotlib with pip install matplotlib")
        return

    economic_label = (
        "LCOE [€/kWh]"
        if ECONOMIC_OBJECTIVE.upper() == "LCOE"
        else "NPC [M€]"
    )
    x_column = _selected_decarbonisation_column()
    x_label = _selected_decarbonisation_label()

    ordered = pareto.sort_values(x_column)
    figure, axis = plt.subplots(figsize=(10, 7), constrained_layout=True)
    axis.plot(
        ordered[x_column],
        ordered["economic_objective"],
        linewidth=1.5,
    )
    points = axis.scatter(
        ordered[x_column],
        ordered["economic_objective"],
        c=ordered["n_wt"],
        cmap="viridis",
        s=75,
        edgecolor="black",
        zorder=2,
    )
    colourbar = figure.colorbar(points, ax=axis)
    colourbar.set_label("Number of installed wind turbines [-]")
    axis.set(
        xlabel=x_label,
        ylabel=economic_label,
        title="Pareto front: economic cost versus decarbonisation",
    )
    axis.grid(alpha=0.3)
    plt.show()


def plot_epsilon_curve(results):
    """Plot minimum economic cost versus achieved decarbonisation."""
    if results.empty or not PLOT_PARETO_FRONT:
        return
    if not MATPLOTLIB_AVAILABLE:
        print("Plot not created: install matplotlib with pip install matplotlib")
        return

    feasible = results[
        results["epsilon_target_feasible"].fillna(False)
    ].copy()
    if feasible.empty:
        print("Epsilon curve not plotted: no feasible target was found.")
        return

    feasible = feasible.sort_values(
        "minimum_decarbonisation_target_percent"
    )

    economic_label = (
        "LCOE [€/kWh]"
        if ECONOMIC_OBJECTIVE.upper() == "LCOE"
        else "NPC [M€]"
    )

    figure, axis = plt.subplots(figsize=(10, 7), constrained_layout=True)
    axis.plot(
        feasible["decarbonisation_percent"],
        feasible["economic_objective"],
        marker="o",
        linewidth=1.6,
    )

    for _, row in feasible.iterrows():
        axis.annotate(
            f"ε={row['minimum_decarbonisation_target_percent']:.0f}%",
            (
                row["decarbonisation_percent"],
                row["economic_objective"],
            ),
            xytext=(4, 5),
            textcoords="offset points",
            fontsize=8,
        )

    axis.set(
        xlabel="Achieved decarbonisation vs diesel-only [%]",
        ylabel=economic_label,
        title="Epsilon-constraint curve: minimum cost vs decarbonisation",
    )
    axis.grid(alpha=0.3)
    plt.show()


def plot_epsilon_design_sizes(results):
    """Plot the optimal component sizes found at each feasible epsilon target."""
    if results.empty or not PLOT_PARETO_DESIGN_SIZES:
        return
    if not MATPLOTLIB_AVAILABLE:
        return

    feasible = results[
        results["epsilon_target_feasible"].fillna(False)
    ].copy()
    if feasible.empty:
        return

    feasible = feasible.sort_values(
        "minimum_decarbonisation_target_percent"
    )

    design_variables = [
        ("pv", "PV size [kW]"),
        ("n_wt", "Number of wind turbines [-]"),
        ("battery", "Battery size [kWh]"),
        ("electrolyzer", "Electrolyzer size [kW]"),
        ("hydrogen_storage", "H2 storage size [kWh_H2]"),
        ("fuel_cell", "Fuel cell size [kW]"),
        ("diesel", "Diesel size [kW]"),
    ]

    figure, axes = plt.subplots(4, 2, figsize=(14, 15), constrained_layout=True)
    x = feasible["minimum_decarbonisation_target_percent"]

    for axis, (column, label) in zip(axes.ravel(), design_variables):
        if column in feasible.columns:
            axis.plot(
                x,
                feasible[column],
                marker="o",
                markersize=4,
                linewidth=1.5,
            )
        axis.set(
            xlabel="Minimum decarbonisation target ε [%]",
            ylabel=label,
        )
        axis.grid(alpha=0.3)

    figure.delaxes(axes.ravel()[-1])
    figure.suptitle(
        "Optimal design sizes along the epsilon-constraint curve",
        fontsize=16,
    )
    plt.show()


def plot_pareto_design_sizes(pareto):
    """Plot component sizes along the selected decarbonisation axis."""
    if not PLOT_PARETO_DESIGN_SIZES or pareto.empty:
        return
    if not MATPLOTLIB_AVAILABLE:
        return

    x_column = _selected_decarbonisation_column()
    x_label = _selected_decarbonisation_label()

    design_variables = [
        ("pv", "PV size [kW]"),
        ("n_wt", "Number of wind turbines [-]"),
        ("battery", "Battery size [kWh]"),
        ("electrolyzer", "Electrolyzer size [kW]"),
        ("hydrogen_storage", "H2 storage size [kWh_H2]"),
        ("fuel_cell", "Fuel cell size [kW]"),
        ("diesel", "Diesel size [kW]"),
    ]

    ordered = pareto.sort_values(x_column)
    figure, axes = plt.subplots(4, 2, figsize=(14, 15), constrained_layout=True)
    for axis, (column, label) in zip(axes.ravel(), design_variables):
        axis.plot(
            ordered[x_column],
            ordered[column],
            marker="o",
            markersize=4,
            linewidth=1.5,
        )
        axis.set(xlabel=x_label, ylabel=label)
        axis.grid(alpha=0.3)

    figure.delaxes(axes.ravel()[-1])
    figure.suptitle("Optimal design sizes along the Pareto front", fontsize=16)
    plt.show()


# =============================================================================
# MAIN
# =============================================================================

if __name__ == "__main__":
    (
        APPLY_HORNSUND_CONSTRAINTS,
        SYSTEM_MODE,
        ECONOMIC_SCENARIO,
        OPTIMIZATION_METHOD,
        EPSILON_MIN_PERCENT,
        EPSILON_MAX_PERCENT,
        EPSILON_STEP_PERCENT,
    ) = select_optimization_options()

    # Explicit sanity check: these are the values that will actually be passed
    # to configure_search_space() and to every candidate evaluation.
    print("\nGUI SELECTION CHECK")
    print(f"  APPLY_HORNSUND_CONSTRAINTS = {APPLY_HORNSUND_CONSTRAINTS}")
    print(f"  SYSTEM_MODE                = {SYSTEM_MODE}")
    print(f"  ECONOMIC_SCENARIO           = {ECONOMIC_SCENARIO}")
    print(f"  OPTIMIZATION_METHOD         = {OPTIMIZATION_METHOD}")

    optimisation_results, model = optimize()

    if OPTIMIZATION_METHOD == "pareto":
        print_pareto_summary(optimisation_results)
        ask_save_results(optimisation_results, model)
        plot_pareto_front(optimisation_results)
        plot_pareto_design_sizes(optimisation_results)
    else:
        print_epsilon_summary(optimisation_results)
        ask_save_results(optimisation_results, model)
        plot_epsilon_curve(optimisation_results)
        plot_epsilon_design_sizes(optimisation_results)
