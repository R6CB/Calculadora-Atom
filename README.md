# Calculadora-Atom
# Calculadora de preus i de rendiment Atom
import os
import math
import requests
import pandas as pd
import numpy as np
import streamlit as st
import openpyxl
import plotly.express as px
import plotly.io as pio
from io import BytesIO
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import sys, importlib.metadata as m, plotly, streamlit as st
with st.expander("🔧 Diagnòstic entorn", expanded=False):
    st.write("Python exe:", sys.executable)
    st.write("Plotly:", plotly.__version__)
    try:
        st.write("Kaleido:", m.version("kaleido"))
        import kaleido
        st.success("Kaleido importat correctament")
    except Exception as e:
        st.error(f"Kaleido NO disponible: {e}")
# ====== PDF (ReportLab) ======
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib.units import mm
from reportlab.lib.pagesizes import A4

# =========================
# Config i costos editables
# =========================
def compute_config_costs(ang:int, N_panels:int, n_li:int, n_h2:int, refills:int,
                         n_fuel_cells:int, costs:dict) -> dict:
    # Parcials (mateixos que al botó 3)
    cost_pan = N_panels * float(costs["PANEL_COST"])
    cost_li  = int(n_li) * float(costs["LI_BATT_UNIT_COST"])
    cost_h2c = int(n_h2) * float(costs["H2_CYL_COST"])
    cost_ref = int(refills) * float(costs["REFILL_COST_PER_CYL"]) * int(n_h2)
    cost_fc  = int(n_fuel_cells) * float(costs["FUEL_CELL_COST"])

    transm_press_cost = float(costs["TRANSMISSOR_PRESS"]) * (int(n_h2) / 2.0)
    valvula_cost      = float(costs["VALVULA"])            * (int(n_h2) / 2.0)

    capex_base = (
        float(costs["CONTAINER"]) + float(costs["RACK"]) + cost_pan + cost_li + cost_h2c + cost_fc +
        float(costs["MPPT"]) + float(costs["CABLE"]) * N_panels + float(costs["REST_SOLAR"]) + float(costs["SAFETY_RELAY"]) + float(costs["DETECTOR_H2"]) +
        float(costs["REST_SAFETY"]) + transm_press_cost + valvula_cost + float(costs["CONTROL_HEATERS"]) +
        float(costs["HEATERS"]) * n_h2 + float(costs["PT100"]) * n_h2 + float(costs["ELECTROLYZER"]) + float(costs["CAUDALIMETRE"]) + float(costs["STEP_DOWN"]) +
        float(costs["ANTENA"]) + float(costs["PLC"]) + float(costs["STEPDOWN_CONTROL"]) + float(costs["STEPDOWN_GENERACIO"]) +
        float(costs["VALVULA_FUELCELL"]) * n_fuel_cells + float(costs["STEPUP"]) * n_fuel_cells +
        float(costs["REDUNDANCEMODULE"]) + float(costs["REDUCTOR_PRESS"]) + float(costs["SENSOR_PRESS"]) + float(costs["CLIMATITZACIO"]) + float(costs["OTHERS"])
    )
    capex = capex_base / (1 - float(costs["MANUFACTORING_%"])) + float(costs["MONTATGE_COST"])
    opex  = cost_ref + float(costs["OPEX_BASE"])
    total = (capex + opex) / (1 - float(costs["MARGE"]))

    return {
        "Angle(°)": int(ang),
        "Plaques": int(N_panels),
        "Bateries_Li": int(n_li),
        "Li_kWh": round(int(n_li) * float(costs["LI_BATT_UNIT_KWH"]), 1),
        "Bombones_H2": int(n_h2),
        "H2_kWh": round(int(n_h2) * float(costs["H2_CYL_KWH"]), 1),
        "Refills": int(refills),
        "FuelCells": int(n_fuel_cells),
        "Cost_plaques(€)": int(cost_pan),
        "Cost_Li(€)": int(cost_li),
        "Cost_bombones(€)": int(cost_h2c),
        "Cost_refills(€)": int(cost_ref),
        "Cost_fuelcells(€)": int(cost_fc),
        "Capex(€)": round(capex, 2),
        "Opex(€)": round(opex, 2),
        "Total(€)": round(total, 2),
    }

#def creadorpdf ():
  #  pdf = canvas.Canvas("Hello world.pdf", pagesize = A4)
  #  h,w = A4
  #  pdf.setFont("Helvetica",10)
  #  pdf.drawString(120, h-50, "ATOM H2")
  #  pdf.drawImage("logo.png", 50, h - 50, width=50, height=50)
  #  pdf.save()

#creadorpdf ()

def default_costs() -> dict:
    return {
        # unitats
        "PANEL_COST": 55.0,
        "LI_BATT_UNIT_KWH": 4.8,
        "LI_BATT_UNIT_COST": 744.0,
        "H2_CYL_KWH": 55.0,
        "H2_CYL_COST": 3500.0,
        "REFILL_COST": 231.2,
        "REFILL_COST_PER_CYL": 57.8,
        "MANUFACTORING_%": 0.30,
        "MARGE": 0.41,

        # CAPEX base / materials
        "CONTAINER": 5634,
        "ESTRUCTURA": 1500,
        "RACK": 1500,
        "MPPT": 392,
        "CABLE": 24,
        "REST_SOLAR": 779,
        "SAFETY_RELAY": 238,
        "DETECTOR_H2": 366,
        "REST_SAFETY": 908,

        "TRANSMISSOR_PRESS": 140,
        "VALVULA": 68,

        "CONTROL_HEATERS": 200,
        "HEATERS": 163,
        "PT100": 40,

        "ELECTROLYZER": 5875,
        "CAUDALIMETRE": 1595,
        "STEP_DOWN": 200,

        "ANTENA": 30,
        "PLC": 3016,
        "STEPDOWN_CONTROL": 100,

        "STEPDOWN_GENERACIO": 146,
        "VALVULA_FUELCELL": 68,
        "STEPUP": 420,
        "REDUNDANCEMODULE": 160,
        "REDUCTOR_PRESS": 164,
        "SENSOR_PRESS": 176,

        "CLIMATITZACIO": 300,
        "OTHERS": 3900,

        # Muntatge / FC
        "MONTATGE_COST": 5000.0,
        "FUEL_CELL_UNIT_KW": 2.0,
        "FUEL_CELL_COST": 5699.0,

        # OPEX extra
        "OPEX_BASE": 25000.0,
        "TOTAL_DIVISOR": 0.59,  # legacy
    }

# =============
# PV i helpers
# =============
def load_pv_from_csv(file) -> pd.DataFrame:
    df = pd.read_csv(file)
    rename = {}
    for c in df.columns:
        lc = c.lower()
        if lc == "datetime": rename[c] = "datetime"
        if lc == "p_kw":     rename[c] = "P_kW"
        if lc in ("p", "power", "p_w"): rename[c] = "P"
    df = df.rename(columns=rename)
    if "datetime" not in df.columns:
        raise ValueError("El CSV ha de tenir una columna 'datetime'.")
    df["datetime"] = pd.to_datetime(df["datetime"], errors="coerce")
    df = df[df["datetime"].notna()].copy()
    if "P_kW" not in df.columns:
        if "P" in df.columns:
            df["P_kW"] = pd.to_numeric(df["P"], errors="coerce") / 1000.0
        else:
            raise ValueError("El CSV ha d’incloure 'P_kW' o 'P' (W).")
    df["Any"]  = df["datetime"].dt.year
    df["Mes"]  = df["datetime"].dt.month
    df["Dia"]  = df["datetime"].dt.day
    df["Hora"] = df["datetime"].dt.hour
    df = df.sort_values("datetime").reset_index(drop=True)
    return df[["Any","Mes","Dia","Hora","P_kW","datetime"]]

PVGIS_URLS = [
    "https://re.jrc.ec.europa.eu/api/v5_3/seriescalc",
    "https://re.jrc.ec.europa.eu/api/v5_2/seriescalc",
    "https://re.jrc.ec.europa.eu/api/seriescalc",
]

def fetch_pvgis(lat, lon, start_year, end_year,
                raddb="PVGIS-SARAH3", pvcalc=1, peak_kw=0.5, loss_pct=14.0,
                angle=30.0, aspect=0.0):
    if pvcalc == 1:
        if peak_kw is None or float(peak_kw) <= 0:
            raise ValueError("La potència pic (kW) ha de ser > 0.")
        if loss_pct is None or not (0 <= float(loss_pct) <= 100):
            raise ValueError("Les pèrdues (%) han d'estar entre 0 i 100.")
    params = {
        "lat": float(lat), "lon": float(lon),
        "raddatabase": raddb,
        "startyear": int(start_year), "endyear": int(end_year),
        "pvcalculation": int(pvcalc),
        "peakpower": float(peak_kw),
        "loss": float(loss_pct),
        "angle": float(angle), "aspect": float(aspect),
        "components": 1, "outputformat": "json", "browser": 0,
    }
    retry = Retry(total=4, connect=4, read=4, backoff_factor=0.8,
                  status_forcelist=(429,500,502,503,504),
                  allowed_methods=frozenset(["GET"]))
    adapter = HTTPAdapter(max_retries=retry)
    s = requests.Session()
    s.mount("https://", adapter); s.mount("http://", adapter)
    last_err = None
    for url in PVGIS_URLS:
        try:
            r = s.get(url, params=params, timeout=30)
            r.raise_for_status()
            js = r.json()
            hourly = js.get("outputs", {}).get("hourly")
            if not hourly:
                raise ValueError("No s'han rebut dades horàries.")
            return pd.json_normalize(hourly)
        except Exception as e:
            last_err = e
            continue
    raise RuntimeError(f"No s'ha pogut contactar amb PVGIS. Detall: {last_err}")

def transform_df(df):
    time_col = next((c for c in df.columns if c.lower().startswith("time")), None)
    if not time_col:
        raise ValueError("No trobo la columna de temps.")
    digits = df[time_col].astype(str).str.replace(r"\D", "", regex=True)
    def split_ymdh(s):
        if len(s) < 10: return pd.Series([None,None,None,None])
        return pd.Series([int(s[0:4]), int(s[4:6]), int(s[6:8]), int(s[8:10])])
    df[["Any","Mes","Dia","Hora"]] = digits.apply(split_ymdh)
    if "P" not in df.columns:  raise ValueError("No trobo 'P' (kW).")
    if "T2m" not in df.columns: df["T2m"] = np.nan
    out = pd.DataFrame({
        "Any": df["Any"], "Mes": df["Mes"], "Dia": df["Dia"], "Hora": df["Hora"],
        "P_kW": pd.to_numeric(df["P"], errors="coerce") / 1000.0,
        "T2m": pd.to_numeric(df["T2m"], errors="coerce"),
    }).sort_values(["Any","Mes","Dia","Hora"]).reset_index(drop=True)
    out["datetime"] = pd.to_datetime(dict(year=out["Any"], month=out["Mes"], day=out["Dia"], hour=out["Hora"]), errors="coerce")
    out = out[out["datetime"].notna()].reset_index(drop=True)
    return out

# ============
# Excel helpers
# ============
def write_new_xlsx_multiple(path, sheet_to_df):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl", mode="w") as w:
        for sheet, df in sheet_to_df.items():
            df.to_excel(w, sheet_name=sheet, index=False)

def overwrite_sheet_in_excel(path, sheet_to_df):
    if not os.path.exists(path):
        raise FileNotFoundError(f"No existeix el fitxer: {path}")
    keep_vba = path.lower().endswith(".xlsm")
    wb = openpyxl.load_workbook(path, keep_vba=keep_vba)
    for sheet, df in sheet_to_df.items():
        if sheet in wb.sheetnames:
            ws = wb[sheet]; ws.delete_rows(1, ws.max_row)
        else:
            ws = wb.create_sheet(sheet)
        for j, col in enumerate(df.columns, start=1):
            ws.cell(row=1, column=j, value=str(col))
        for i, row in enumerate(df.itertuples(index=False), start=2):
            for j, val in enumerate(row, start=1):
                ws.cell(row=i, column=j, value=val)
    wb.save(path)

# ============
# Simulacions
# ============
def simulate_battery(df, cap_kwh, soc_ini, soc_min, p_ch_max, p_dis_max, eta_c, eta_d, load_kw_fixed):
    dff = df.copy().sort_values("datetime").reset_index(drop=True)
    diffs = dff["datetime"].diff().dropna()
    dt_hours = diffs.dt.total_seconds().mode().iloc[0] / 3600.0 if len(diffs) else 1.0
    n = len(dff)
    soc = np.zeros(n); soc[0] = float(soc_ini)
    imp = np.zeros(n); exp = np.zeros(n)
    p_ch = np.zeros(n); p_dis = np.zeros(n)
    dff["load_kw"] = float(load_kw_fixed)
    for i in range(n):
        gen = float(pd.to_numeric(dff["P_kW"].iat[i], errors="coerce") or 0.0)
        load = float(dff["load_kw"].iat[i])
        surplus = gen - load
        prev_soc = soc[i-1] if i > 0 else soc[0]
        if surplus >= 0:
            p_avail = min(surplus, float(p_ch_max))
            e_to_bat = p_avail * dt_hours * float(eta_c)
            e_room = max(0.0, (1.0 - prev_soc) * float(cap_kwh))
            e_store = min(e_to_bat, e_room)
            new_soc = prev_soc + (e_store / float(cap_kwh) if float(cap_kwh) > 0 else 0.0)
            soc[i] = np.clip(new_soc, 0.0, 1.0)
            p_ch[i] = (e_store / dt_hours) / float(eta_c) if dt_hours > 0 and float(eta_c) > 0 else 0.0
            exp[i] = max(0.0, surplus - p_ch[i]); imp[i] = 0.0
        else:
            demand = -surplus
            p_from_bat = min(float(p_dis_max), demand)
            e_from_bat = p_from_bat * dt_hours / (float(eta_d) if float(eta_d) > 0 else 1.0)
            e_avail = max(0.0, (prev_soc - float(soc_min)) * float(cap_kwh))
            e_draw = min(e_from_bat, e_avail)
            new_soc = prev_soc - (e_draw / float(cap_kwh) if float(cap_kwh) > 0 else 0.0)
            soc[i] = np.clip(new_soc, float(soc_min), 1.0)
            p_dis[i] = (e_draw * (float(eta_d) if float(eta_d) > 0 else 1.0)) / dt_hours if dt_hours > 0 else 0.0
            imp[i] = max(0.0, demand - p_dis[i]); exp[i] = 0.0
    dff["pv_kw"] = dff["P_kW"]
    dff["import_kw"] = imp; dff["export_kw"] = exp
    dff["bat_charge_kw"] = p_ch; dff["bat_discharge_kw"] = p_dis
    dff["soc_%"] = soc * 100.0
    return dff, dt_hours

def simulate_h2(df_hourly: pd.DataFrame, dt_hours: float,
                cap_h2_kwh: float, soc_h2_init_pct: float,
                eta_charge: float = 0.60, eta_discharge: float = 0.40,
                load_kw_fixed: float = 2.0, soc_li_min_frac: float = 0.10,
                refill_threshold_frac: float = 0.15):
    v = df_hourly.copy().reset_index(drop=True)
    v.columns = [str(c).strip() for c in v.columns]
    for col in ["import_kw", "export_kw"]:
        if col not in v.columns:
            v[col] = 0.0
    if "soc_%" not in v.columns:
        v["soc_%"] = soc_li_min_frac * 100.0
    n = len(v); cap = float(cap_h2_kwh)
    h2_soc = np.zeros(n); h2_charge = np.zeros(n)
    h2_draw = np.zeros(n); h2_deliv = np.zeros(n)
    imp_after_kw = np.array(pd.to_numeric(v["import_kw"], errors="coerce").fillna(0.0), dtype=float)
    refill_event = np.zeros(n, dtype=int)
    h2_soc[0] = cap * float(soc_h2_init_pct)
    refill_next = False; eps = 1e-9
    for i in range(n):
        soc_prev = h2_soc[i-1] if i > 0 else h2_soc[0]
        if refill_next:
            soc_prev = cap; refill_event[i] = 1; refill_next = False
        exp_kw = float(pd.to_numeric(v["export_kw"].iat[i], errors="coerce") or 0.0)
        e_export = exp_kw * dt_hours
        e_store_eff = e_export * float(eta_charge)
        room = max(0.0, cap - soc_prev)
        e_store = min(e_store_eff, room)
        soc_after_charge = soc_prev + e_store
        h2_charge[i] = e_store
        imp_kw = float(pd.to_numeric(v["import_kw"].iat[i], errors="coerce") or 0.0)
        e_import = imp_kw * dt_hours
        soc_li_pct = float(pd.to_numeric(v["soc_%"].iat[i], errors="coerce") or 0.0)
        if e_import > 0 and soc_li_pct <= (100.0 * soc_li_min_frac + 1e-6):
            e_draw_wish = 2.0 * float(load_kw_fixed) * dt_hours
            e_draw = min(e_draw_wish, soc_after_charge)
            e_deliv = e_draw * float(eta_discharge)
            e_import_new = max(0.0, e_import - e_deliv)
            imp_after_kw[i] = (e_import_new / dt_hours) if dt_hours > 0 else imp_after_kw[i]
            h2_draw[i] = e_draw; h2_deliv[i] = e_deliv
            soc_end = soc_after_charge - e_draw
        else:
            h2_draw[i] = 0.0; h2_deliv[i] = 0.0; soc_end = soc_after_charge
        h2_soc[i] = soc_end
        if h2_soc[i] < cap * float(refill_threshold_frac) - eps:
            refill_next = True
    v["h2_soc_kwh"] = h2_soc
    v["h2_charge_kwh"] = h2_charge
    v["h2_draw_kwh"] = h2_draw
    v["h2_delivered_kwh"] = h2_deliv
    v["import_after_h2_kw"] = imp_after_kw
    v["h2_refill_event"] = refill_event
    v["h2_refills_cum"] = np.cumsum(refill_event)
    return v

# =========
# Packing
# =========
def layout_panels_for_angle(L,W,panel_L,panel_W,theta_deg,margin=0.0,g_long=0.0,g_lat=0.0,
                            orientation="portrait",avoid_shade=False,alpha_min_deg=20.0,
                            optimize_row_axis=True):
    L_eff = max(0.0, L - 2*margin); W_eff = max(0.0, W - 2*margin)
    if L_eff <= 0 or W_eff <= 0: return 0,0,0,None,None,None
    if orientation == "portrait":
        mod_long = panel_L; mod_wide = panel_W
    else:
        mod_long = panel_W; mod_wide = panel_L
    theta = math.radians(theta_deg)
    long_proj = mod_long * math.cos(theta)
    h = mod_long * math.sin(theta)
    if avoid_shade:
        alpha = max(1e-3, math.radians(alpha_min_deg))
        s_shade = h / math.tan(alpha)
        step_long = long_proj + max(g_long, s_shade)
    else:
        step_long = long_proj + g_long
    step_lat = mod_wide + g_lat
    def pack(L_rows_axis, W_cols_axis):
        n_rows = int(L_rows_axis // step_long) if step_long > 0 else 0
        n_cols = int(W_cols_axis // step_lat) if step_lat > 0 else 0
        return n_rows, n_cols, n_rows*n_cols
    n_rows_A, n_cols_A, N_A = pack(L_eff, W_eff)
    n_rows_B, n_cols_B, N_B = pack(W_eff, L_eff)
    if optimize_row_axis and N_B > N_A:
        return N_B, n_rows_B, n_cols_B, "rows//W", step_long, step_lat
    else:
        return N_A, n_rows_A, n_cols_A, "rows//L", step_long, step_lat

# =========
# UI
# =========
st.set_page_config(page_title="Calculadora Atom", layout="wide")
st.title("Calculadora d'eficiència i costos ATOM H2")

if "xrange" not in st.session_state: st.session_state.xrange = None
if "layout_df" not in st.session_state: st.session_state.layout_df = None
if "costs" not in st.session_state: st.session_state.costs = default_costs()

# 1) Formulari de paràmetres (inclou PREUS I COSTOS a dins)
with st.form("params_form"):
    st.subheader("Angles d'inclinació")
    all_angles = list(range(20, 61, 5))
    angles = st.multiselect("Selecciona angles (°)", options=all_angles, default=[40], key="pv_angles")
    if not angles:
        st.caption("Selecciona almenys un angle.")

    with st.expander("Paràmetres avançats — PV i Excel", expanded=False):
        lat = st.number_input("Latitud", value=41.355, step=0.001, format="%.6f", key="pv_lat")
        lon = st.number_input("Longitud", value=1.852, step=0.001, format="%.6f", key="pv_lon")
        start_year = st.number_input("Any inici", value=2010, step=1, key="pv_sy")
        end_year = st.number_input("Any fi", value=2020, step=1, key="pv_ey")
        aspect = st.number_input("Azimut (°)", value=0.0, step=1.0, key="pv_az")
        peak_kw = st.number_input("Pot. pic d'UN panell (kW)", value=0.5, min_value=0.001, step=0.1, key="pv_peak")
        loss_pct = st.number_input("Pèrdues sistema (%)", value=14.0, min_value=0.0, max_value=100.0, step=0.5, key="pv_loss")
        pv_source = st.selectbox("Origen de dades PV", ["PVGIS (online)", "CSV local (offline)"], index=0, key="pv_src")
        pv_csv = None
        if pv_source == "CSV local (offline)":
            pv_csv = st.file_uploader("Puja CSV (datetime, P_kW per UN panell; o datetime, P en W)", type=["csv"], key="pv_csv")

        default_path = os.path.join(os.path.expanduser("~"), "Desktop", "PVGIS_Data.xlsx")
        path = st.text_input("Ruta arxiu de sortida (.xlsx o .xlsm)", value=default_path, key="pv_path")
        base_sheet = st.text_input("Nom base del full", value="Data", key="pv_sheet")
        mode = st.radio("Acció a Excel", ["Crear Excel nou (.xlsx)", "Modificar existent (crear/sobreescriure fulls)"], index=0, key="pv_mode")

    # --- PREUS I COSTOS (DINS DEL FORM)
    with st.expander("Preus i costos", expanded=False):
        st.caption("Ajusta els preus i costos. Es guarden només mentre dura la sessió.")
        costs = st.session_state.costs
        colA, colB, colC = st.columns(3)
        with colA:
            costs["PANEL_COST"] = st.number_input("Preu panell (€/u)", value=float(costs["PANEL_COST"]), step=1.0, key="c_panel")
            costs["LI_BATT_UNIT_COST"] = st.number_input("Bateria Li (€/u)", value=float(costs["LI_BATT_UNIT_COST"]), step=1.0, key="c_li_cost")
            costs["LI_BATT_UNIT_KWH"] = st.number_input("Bateria Li (kWh/u)", value=float(costs["LI_BATT_UNIT_KWH"]), step=0.1, key="c_li_kwh")
            costs["H2_CYL_COST"] = st.number_input("Bombona H₂ (€/u)", value=float(costs["H2_CYL_COST"]), step=1.0, key="c_h2_cost")
            costs["H2_CYL_KWH"] = st.number_input("Bombona H₂ (kWh/u)", value=float(costs["H2_CYL_KWH"]), step=1.0, key="c_h2_kwh")
            costs["REFILL_COST"] = st.number_input("Cost refill H₂ (€/refill)", value=float(costs["REFILL_COST"]), step=1.0, key="c_refill")
            costs["REFILL_COST_PER_CYL"] = st.number_input("Cost refill per bombona (€)", value=float(costs["REFILL_COST_PER_CYL"]), step=0.1, key="c_refill_cyl")
            costs["MANUFACTORING_%"] = st.number_input("Manufacturing (fracció)", value=float(costs["MANUFACTORING_%"]), step=0.01, min_value=0.0, max_value=0.99, key="c_manu")

        with colB:
            costs["CONTAINER"] = st.number_input("Container (€)", value=float(costs["CONTAINER"]), step=10.0, key="c_cont")
            costs["ESTRUCTURA"] = st.number_input("Estructura (€)", value=float(costs["ESTRUCTURA"]), step=10.0, key="c_estr")
            costs["RACK"] = st.number_input("Rack (€)", value=float(costs["RACK"]), step=10.0, key="c_rack")
            costs["MPPT"] = st.number_input("MPPT (€)", value=float(costs["MPPT"]), step=1.0, key="c_mppt")
            costs["CABLE"] = st.number_input("Cable (€/panell)", value=float(costs["CABLE"]), step=1.0, key="c_cable")
            costs["REST_SOLAR"] = st.number_input("Rest solar (€)", value=float(costs["REST_SOLAR"]), step=1.0, key="c_restsolar")
            costs["SAFETY_RELAY"] = st.number_input("Safety relay (€)", value=float(costs["SAFETY_RELAY"]), step=1.0, key="c_safety")
            costs["DETECTOR_H2"] = st.number_input("Detector H₂ (€)", value=float(costs["DETECTOR_H2"]), step=1.0, key="c_detect")
            costs["REST_SAFETY"] = st.number_input("Rest safety (€)", value=float(costs["REST_SAFETY"]), step=1.0, key="c_restsafe")
        with colC:
            costs["TRANSMISSOR_PRESS"] = st.number_input("Transmissor pressió (€)", value=float(costs["TRANSMISSOR_PRESS"]), step=1.0, key="c_tpress")
            costs["VALVULA"] = st.number_input("Vàlvula (€)", value=float(costs["VALVULA"]), step=1.0, key="c_valv")
            costs["CONTROL_HEATERS"] = st.number_input("Control heaters (€)", value=float(costs["CONTROL_HEATERS"]), step=1.0, key="c_ctrlheat")
            costs["HEATERS"] = st.number_input("Heaters (€/bombona)", value=float(costs["HEATERS"]), step=1.0, key="c_heaters")
            costs["PT100"] = st.number_input("PT100 (€/bombona)", value=float(costs["PT100"]), step=1.0, key="c_pt100")
            costs["ELECTROLYZER"] = st.number_input("Electrolitzador (€)", value=float(costs["ELECTROLYZER"]), step=1.0, key="c_elec")
            costs["CAUDALIMETRE"] = st.number_input("Caudalímetre (€)", value=float(costs["CAUDALIMETRE"]), step=1.0, key="c_caudal")
            costs["STEP_DOWN"] = st.number_input("Step-down (€)", value=float(costs["STEP_DOWN"]), step=1.0, key="c_stepdown")
            costs["ANTENA"] = st.number_input("Antena (€)", value=float(costs["ANTENA"]), step=1.0, key="c_ant")
            costs["PLC"] = st.number_input("PLC (€)", value=float(costs["PLC"]), step=1.0, key="c_plc")
            costs["STEPDOWN_CONTROL"] = st.number_input("Step-down control (€)", value=float(costs["STEPDOWN_CONTROL"]), step=1.0, key="c_sdc")
            costs["STEPDOWN_GENERACIO"] = st.number_input("Step-down generació (€)", value=float(costs["STEPDOWN_GENERACIO"]), step=1.0, key="c_sdg")
            costs["VALVULA_FUELCELL"] = st.number_input("Vàlvula FC (€/u FC)", value=float(costs["VALVULA_FUELCELL"]), step=1.0, key="c_vfc")
            costs["STEPUP"] = st.number_input("Step-up (€/u FC)", value=float(costs["STEPUP"]), step=1.0, key="c_stepup")
            costs["REDUNDANCEMODULE"] = st.number_input("Mòdul redundància (€)", value=float(costs["REDUNDANCEMODULE"]), step=1.0, key="c_red")
            costs["REDUCTOR_PRESS"] = st.number_input("Reductor pressió (€)", value=float(costs["REDUCTOR_PRESS"]), step=1.0, key="c_reductor")
            costs["SENSOR_PRESS"] = st.number_input("Sensor pressió (€)", value=float(costs["SENSOR_PRESS"]), step=1.0, key="c_sensor")
            costs["CLIMATITZACIO"] = st.number_input("Climatització (€)", value=float(costs["CLIMATITZACIO"]), step=1.0, key="c_clima")
            costs["OTHERS"] = st.number_input("Altres (€)", value=float(costs["OTHERS"]), step=1.0, key="c_others")

        st.markdown("---")
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1:
            costs["MONTATGE_COST"] = st.number_input("Muntatge (CAPEX) (€)", value=float(costs["MONTATGE_COST"]), step=10.0, key="c_mont")
        with col2:
            costs["FUEL_CELL_UNIT_KW"] = st.number_input("Pila combustible (kW/u)", value=float(costs["FUEL_CELL_UNIT_KW"]), step=0.1, key="c_fc_kw")
        with col3:
            costs["FUEL_CELL_COST"] = st.number_input("Pila combustible (€/u)", value=float(costs["FUEL_CELL_COST"]), step=10.0, key="c_fc_cost")
        with col4:
            costs["OPEX_BASE"] = st.number_input("PREU_OPERARTIU (€)", value=float(costs["OPEX_BASE"]), step=100.0, key="c_opex")
        with col5:
            costs["MARGE"] = st.number_input("Marge (fracció)", value=float(costs["MARGE"]), step=0.01, min_value=0.0, max_value=0.99, key="c_margin")

    # --- Paràmetres bateria i càrrega (Lítio)
    with st.expander("Paràmetres bateria i càrrega (Lítio)", expanded=False):
        n_batt_li = st.number_input("Nombre de bateries Li (4.8 kWh/u)", value=10, min_value=0, step=1, key="li_nbatt")
        cap_kwh = n_batt_li * float(st.session_state.costs["LI_BATT_UNIT_KWH"])
        st.caption(f"Capacitat total: {cap_kwh:.1f} kWh")
        soc_ini = st.slider("SOC inicial (%)", 0, 100, 100, 1, key="li_soc_ini") / 100.0
        soc_min = st.slider("SOC mínim (%)", 0, 50, 20, 1, key="li_soc_min") / 100.0
        p_ch_max = st.number_input("Pot. màx. càrrega (kW)", value=21.0, min_value=0.0, step=0.1, key="li_pch")
        p_dis_max = st.number_input("Pot. màx. descàrrega (kW)", value=5.0, min_value=0.0, step=0.1, key="li_pdis")
        eta_c = st.slider("Rendiment càrrega", 0.5, 1.0, 0.95, 0.01, key="li_eta_c")
        eta_d = st.slider("Rendiment descàrrega", 0.5, 1.0, 0.95, 0.01, key="li_eta_d")
        load_kw_fixed = st.number_input("Consum fix (kW)", value=2.0, min_value=0.0, step=0.1, key="li_load")
        p_max_kw = st.number_input("Consum màxim (kW)", value=2.0, min_value=0.0, step=0.1, key="li_pmax")
        n_fuel_cells_global = int(math.ceil(p_max_kw / float(st.session_state.costs["FUEL_CELL_UNIT_KW"]))) if p_max_kw > 0 else 0
        st.caption(f"Piles de combustible requerides ({st.session_state.costs['FUEL_CELL_UNIT_KW']} kW/u): {n_fuel_cells_global} u — cost {n_fuel_cells_global*st.session_state.costs['FUEL_CELL_COST']:,.0f} €")

    # --- Bateria d'hidrogen
    with st.expander("Bateria d'hidrogen (H₂)", expanded=False):
        cap_h2_kwh = st.number_input("Capacitat H₂ (kWh)", value=220.0, min_value=0.0, step=0.5, key="h2_cap")
        soc_h2_init_pct = st.slider("SOC inicial H₂ (%)", 0, 100, 100, 1, key="h2_soc_ini") / 100.0
        eta_h2_charge = st.slider("Eficiència càrrega H₂", 0.0, 1.0, 0.60, 0.01, key="h2_eta_c")
        eta_h2_discharge = st.slider("Eficiència descàrrega H₂", 0.0, 1.0, 0.40, 0.01, key="h2_eta_d")
        refill_threshold_pct = st.slider("Llindar refill H₂ (%)", 5, 50, 15, 1, key="h2_thr")

    # --- Parcel·la
    with st.expander("Parcela, mòduls i empaquetat", expanded=False):
        st.markdown("**Dimensions de la parcel·la (m)**")
        L = st.number_input("Llargada L (m)", value=10.0, min_value=0.0, step=0.1, key="pack_L")
        W = st.number_input("Amplada W  (m)", value=10.0, min_value=0.0, step=0.1, key="pack_W")
        margin = st.number_input("Marge perimetral (m)", value=0.0, min_value=0.0, step=0.05, key="pack_margin")
        st.markdown("**Dimensions del panell (m)**")
        panel_L = st.number_input("Costat llarg del mòdul (m)", value=2.0, min_value=0.1, step=0.05, key="pack_panelL")
        panel_W = st.number_input("Costat ample del mòdul (m)", value=1.0, min_value=0.1, step=0.05, key="pack_panelW")
        orientation = st.selectbox("Orientació del mòdul", ["portrait", "landscape"], index=0, key="pack_or")
        st.markdown("**Espais entre mòduls (m)**")
        g_long = st.number_input("Espai longitudinal (entre files)", value=0.0, min_value=0.0, step=0.05, key="pack_glong")
        g_lat  = st.number_input("Espai lateral (entre columnes)",  value=0.0, min_value=0.0, step=0.05, key="pack_glat")
        st.markdown("**Estratègia d'ombres**")
        avoid_shade = st.checkbox("Evitar auto-ombres (conservador)", value=False, key="pack_shade")
        alpha_min_deg = st.number_input("Alçada solar mínima α (°)", value=20.0, min_value=1.0, max_value=89.0, step=1.0, key="pack_alpha")
        optimize_row_axis = st.checkbox("Optimitza eix de files (provar //L i //W)", value=True, key="pack_opt")

    # --- Botons del FORM
    bcol1, bcol2, bcol3 = st.columns([1, 2, 2])
    with bcol1:
        reset_costs = st.form_submit_button("🔄 Restableix costos", use_container_width=True)
    with bcol2:
        calc = st.form_submit_button("1) Calcula plaques per angle", use_container_width=True)
    with bcol3:
        run = st.form_submit_button("2) Run descarrega + simulació", use_container_width=True)

# --- Després del form: gestió del botó de restablir costos
if reset_costs:
    st.session_state.costs = default_costs()
    st.success("Costos restablerts als valors per defecte.")
    st.rerun()

# Assegura costs actualitzat
costs = st.session_state.costs

def resample_for_view(ddf: pd.DataFrame, gran: str):
    ddf = ddf.set_index("datetime")
    if gran == "Horària": return ddf
    agg = {"pv_kw":"mean","load_kw":"mean","import_kw":"mean","export_kw":"mean",
           "bat_charge_kw":"mean","bat_discharge_kw":"mean","soc_%":"last"}
    rule = "D" if gran == "Diària" else "MS"
    return ddf.resample(rule).agg(agg)

# ---------- BOTÓ 1 ----------
if calc:
    try:
        rows = []
        for ang in angles:
            N, n_rows, n_cols, best_axis, step_long, step_lat = layout_panels_for_angle(
                L=L, W=W, panel_L=panel_L, panel_W=panel_W, theta_deg=float(ang),
                margin=margin, g_long=g_long, g_lat=g_lat,
                orientation=orientation, avoid_shade=avoid_shade,
                alpha_min_deg=alpha_min_deg, optimize_row_axis=optimize_row_axis
            )
            rows.append({
                "Angle(°)": int(ang), "N_panells": int(N), "Files": int(n_rows), "Columnes": int(n_cols),
                "Files_eix": best_axis, "pas_long(m)": None if step_long is None else round(step_long,3),
                "pas_lat(m)": None if step_lat is None else round(step_lat,3),
                "kWp_total": round(N * float(peak_kw), 3)
            })
        layout_df = pd.DataFrame(rows).sort_values("Angle(°)").reset_index(drop=True)
        st.session_state.layout_df = layout_df
        st.success("Packing calculat i desat.")
        st.subheader("Packing per angle (previsió)")
        st.dataframe(layout_df)
    except Exception as e:
        st.error(f"Error al càlcul del packing: {e}")

# ---------- BOTÓ 2 ----------
if run:
    try:
        layout_df = st.session_state.layout_df
        if layout_df is None:
            rows = []
            for ang in angles:
                N, n_rows, n_cols, best_axis, step_long, step_lat = layout_panels_for_angle(
                    L=L, W=W, panel_L=panel_L, panel_W=panel_W, theta_deg=float(ang),
                    margin=margin, g_long=g_long, g_lat=g_lat, orientation=orientation,
                    avoid_shade=avoid_shade, alpha_min_deg=alpha_min_deg, optimize_row_axis=optimize_row_axis
                )
                rows.append({"Angle(°)": int(ang), "N_panells": int(N), "kWp_total": round(N * float(peak_kw), 3)})
            layout_df = pd.DataFrame(rows).sort_values("Angle(°)").reset_index(drop=True)
            st.session_state.layout_df = layout_df
            st.info("No hi havia packing previ. L'he calculat amb els paràmetres actuals.")

        valid = layout_df[layout_df["N_panells"] > 0].copy()
        if valid.empty:
            st.error("Cap angle amb N>0. Revisa paràmetres de parcel·la/espais.")
        else:
            sheet_to_df = {}; views_soc=[]; views_pv=[]; flows_all=[]; views_h2soc=[]
            refills_summary=[]; h2_marks=[]
            for _, row in valid.iterrows():
                ang = int(row["Angle(°)"]); N_panels = int(row["N_panells"])
                # Obté sèrie PV per UN panell
                if pv_source == "CSV local (offline)":
                    if pv_csv is None:
                        st.error("Selecciona un CSV per al mode offline."); st.stop()
                    base_one = load_pv_from_csv(pv_csv)
                else:
                    raw = fetch_pvgis(lat=lat, lon=lon, start_year=int(start_year), end_year=int(end_year),
                                      raddb="PVGIS-SARAH3", pvcalc=1, peak_kw=float(peak_kw), loss_pct=float(loss_pct),
                                      angle=float(ang), aspect=float(aspect))
                    base_one = transform_df(raw)
                base = base_one.copy(); base["P_kW"] = base["P_kW"] * N_panels
                export_df = base[["Any","Mes","Dia","Hora","P_kW","T2m","datetime"]].copy()
                sheet_name = f"{base_sheet}_{int(ang):02d}deg" if len(valid) > 1 else base_sheet
                sheet_to_df[sheet_name] = export_df

                # Simula lítio
                sim, _dt = simulate_battery(df=base, cap_kwh=cap_kwh, soc_ini=soc_ini, soc_min=soc_min,
                                            p_ch_max=p_ch_max, p_dis_max=p_dis_max, eta_c=eta_c, eta_d=eta_d,
                                            load_kw_fixed=load_kw_fixed)
                v = resample_for_view(sim.copy(), gran="Horària").reset_index()
                v["Angle"] = f"{int(ang)}°"
                views_soc.append(v[["datetime","soc_%","Angle"]])
                views_pv.append(v[["datetime","pv_kw","Angle"]])

                if "soc_%" not in v.columns: v["soc_%"] = soc_min * 100.0

                # Simula H2
                vh2 = simulate_h2(df_hourly=v, dt_hours=_dt, cap_h2_kwh=cap_h2_kwh,
                                  soc_h2_init_pct=soc_h2_init_pct, eta_charge=eta_h2_charge,
                                  eta_discharge=eta_h2_discharge, load_kw_fixed=load_kw_fixed,
                                  soc_li_min_frac=soc_min, refill_threshold_frac=(refill_threshold_pct/100.0))
                vh2["h2_soc_%"] = np.where(cap_h2_kwh > 0, 100.0 * vh2["h2_soc_kwh"] / cap_h2_kwh, 0.0)
                refills_summary.append({"Angle": f"{int(ang)}°", "Refills": int(vh2["h2_refill_event"].sum())})
                if "h2_refill_event" in vh2.columns:
                    dts = vh2.loc[vh2["h2_refill_event"] == 1, "datetime"]
                    for dt in dts: h2_marks.append({"datetime": dt, "Angle": f"{int(ang)}°"})
                vh2["Angle"] = f"{int(ang)}°"; views_h2soc.append(vh2[["datetime","h2_soc_%","Angle"]])

                # Fluxos
                flows_cols = ["pv_kw","load_kw","import_kw","export_kw","bat_charge_kw","bat_discharge_kw"]
                vf = v[["datetime","Angle"] + [c for c in flows_cols if c in v.columns]].copy()
                if "import_after_h2_kw" in vh2.columns:
                    vf = vf.merge(vh2[["datetime","import_after_h2_kw"]], on="datetime", how="left")
                melted = vf.melt(id_vars=["datetime","Angle"], var_name="Sèrie", value_name="kW")
                flows_all.append(melted)

            # Escriure Excel (de dades PV simulades per cada angle)
            if mode.startswith("Crear"):
                if not path.lower().endswith(".xlsx"): path = os.path.splitext(path)[0] + ".xlsx"
                write_new_xlsx_multiple(path, sheet_to_df)
            else:
                overwrite_sheet_in_excel(path, sheet_to_df)
            st.success(f"Escrit a «{path}» ({len(sheet_to_df)} fulls).")

            # ------ Costos per angle (mateix format que el botó 3) ------
            if refills_summary:
                detailed_rows = []
                ref_df = pd.DataFrame(refills_summary)  # {'Angle': '40°', 'Refills': X}
                ref_df["Angle(°)"] = ref_df["Angle"].str.replace("°","", regex=False).astype(int)

                for _, r in valid.iterrows():
                    ang = int(r["Angle(°)"])
                    N_panels = int(r["N_panells"])
                    n_li = int(st.session_state["li_nbatt"])
                    n_h2 = max(1, int(math.ceil(cap_h2_kwh / float(costs["H2_CYL_KWH"]))))
                    n_fc = int(n_fuel_cells_global)
                    refills = int(ref_df.loc[ref_df["Angle(°)"] == ang, "Refills"].sum())

                    detailed_rows.append(
                        compute_config_costs(
                            ang=ang, N_panels=N_panels, n_li=n_li, n_h2=n_h2,
                            refills=refills, n_fuel_cells=n_fc, costs=costs
                        )
                    )

                ref_full = pd.DataFrame(detailed_rows).sort_values("Total(€)").reset_index(drop=True)
                st.subheader("Costos per angle (mateix format que l'Excel del botó 3)")
                st.dataframe(ref_full)

                c1, c2 = st.columns(2)
                with c1:
                    csv_bytes = ref_full.to_csv(index=False).encode("utf-8")
                    st.download_button("Descarrega resum (CSV)", data=csv_bytes,
                                       file_name="resum_costos_angles.csv", mime="text/csv")
                with c2:
                    xls_buf = BytesIO()
                    with pd.ExcelWriter(xls_buf, engine="openpyxl") as writer:
                        ref_full.to_excel(writer, sheet_name="Costos", index=False)
                    st.download_button("Descarrega resum (XLSX)", data=xls_buf.getvalue(),
                                       file_name="resum_costos_angles.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            # Gràfics
            soc_df = pd.concat(views_soc, ignore_index=True)
            pv_df  = pd.concat(views_pv,  ignore_index=True)
            h2_df  = pd.concat(views_h2soc, ignore_index=True)

            st.subheader("SOC de la bateria (Lítio) — comparació per angle")
            fig_soc = px.line(soc_df, x="datetime", y="soc_%", color="Angle",
                              labels={"datetime": "Temps", "soc_%": "SOC (%)"})
            fig_soc.update_xaxes(rangeslider=dict(visible=True))
            st.plotly_chart(fig_soc, use_container_width=True)

            st.subheader("Generació PV (kW) — comparació per angle")
            fig_pv = px.line(pv_df, x="datetime", y="pv_kw", color="Angle",
                             labels={"datetime": "Temps", "pv_kw": "kW"})
            fig_pv.update_xaxes(rangeslider=dict(visible=True))
            st.plotly_chart(fig_pv, use_container_width=True)

            st.subheader("Fluxos de potència (kW) — comparació per angle (inclou import_after_h2)")
            flows_df = pd.concat(flows_all, ignore_index=True) if flows_all else pd.DataFrame()
            if not flows_df.empty:
                fig_flow = px.line(flows_df, x="datetime", y="kW", color="Sèrie", line_dash="Angle",
                                   labels={"datetime": "Temps", "kW": "kW"})
                fig_flow.update_xaxes(rangeslider=dict(visible=True))
                st.plotly_chart(fig_flow, use_container_width=True)
            else:
                st.info("No hi ha dades de fluxos disponibles per pintar.")

            st.subheader("SOC bateria d'hidrogen (H₂) — comparació per angle")
            fig_h2 = px.line(h2_df, x="datetime", y="h2_soc_%", color="Angle",
                             labels={"datetime": "Temps", "h2_soc_%": "SOC H₂ (%)"})
            fig_h2.update_xaxes(rangeslider=dict(visible=True))
            st.plotly_chart(fig_h2, use_container_width=True)

    except Exception as e:
        st.error(f"S'ha produït un error: {e}")

elif st.session_state.layout_df is not None:
    st.subheader("Packing per angle (últim càlcul)")
    st.dataframe(st.session_state.layout_df)

# =======================
# ===== Helpers PDF =====
# =======================
def fig_to_png_bytes(fig, scale=2.0, width=1000, height=500):
    """
    Converteix una figura Plotly a PNG (bytes).
    Necessita plotly>=6.1.1 o bé kaleido==0.2.1 si uses plotly 5.x.
    """
    return pio.to_image(fig, format="png", width=width, height=height, scale=scale)

def draw_title(c, text, x, y, size=16, color=(0,0,0)):
    c.setFillColorRGB(*color); c.setFont("Helvetica-Bold", size); c.drawString(x, y, text)

def draw_label(c, text, x, y, size=10, color=(0,0,0)):
    c.setFillColorRGB(*color); c.setFont("Helvetica", size); c.drawString(x, y, text)

def draw_kpi_box(c, x, y, w, h, title, value):
    c.setStrokeColor(colors.HexColor("#E5E7EB"))
    c.setFillColor(colors.white)
    c.roundRect(x, y, w, h, 6, stroke=1, fill=1)
    draw_label(c, title, x+8, y+h-16, size=9, color=(0.3,0.3,0.3))
    draw_title(c, value, x+8, y+h-36, size=14, color=(0,0,0))

def draw_image_fullwidth(c, png_bytes, x, y, max_w):
    img = ImageReader(BytesIO(png_bytes))
    iw, ih = img.getSize()
    scale = max_w / iw
    new_w = max_w
    new_h = ih * scale
    c.drawImage(img, x, y, width=new_w, height=new_h, preserveAspectRatio=True, mask='auto')
    return new_h

def build_pdf_report(
    pdf_bytes_io: BytesIO,
    kpis: dict,
    location_tuple: tuple,
    period_str: str,
    angles_list: list,
    base_load_kw: float,
    fuel_cell_kw_unit: float,
    recommended_angle: int,
    figs: dict,
    costs_table_df: pd.DataFrame
):
    c = canvas.Canvas(pdf_bytes_io, pagesize=A4)
    W, H = A4
    margin = 15*mm
    x0 = margin
    y = H - margin

    # Capçalera
    draw_title(c, "Informe tècnic — Sistema PV + Li + H₂", x0, y, size=18); y -= 10*mm
    lat, lon = location_tuple
    draw_label(c, f"Ubicació: {lat}, {lon}   ·   Període: {period_str}", x0, y); y -= 6*mm
    draw_label(c, f"Angles: {', '.join([str(a)+'°' for a in angles_list])}   ·   Consum base: {base_load_kw} kW   ·   Piles de combustible: {fuel_cell_kw_unit} kW/u", x0, y); y -= 8*mm

    # KPIs
    box_w = (W - 2*margin - 3*6*mm) / 4.0
    box_h = 22*mm
    x = x0
    kpi_items = [
        ("kWh PV total", f"{kpis.get('pv_total_mwh',0):.2f} MWh"),
        ("Export neta", f"{kpis.get('export_mwh',0):.2f} MWh"),
        ("Import post-H₂", f"{kpis.get('import_after_h2_mwh',0):.2f} MWh"),
        ("Refills H₂/any", f"{kpis.get('refills',0)}"),
    ]
    for title, val in kpi_items:
        draw_kpi_box(c, x, y-box_h, box_w, box_h, title, val)
        x += box_w + 6*mm
    y -= (box_h + 10*mm)

    # Recomanació
    draw_label(c, "Recomanació ràpida", x0, y, size=10, color=(0.25,0.25,0.25)); y -= 6*mm
    draw_title(c, f"Angle recomanat: {recommended_angle}° pel millor cost total estimat", x0, y, size=12, color=(0,0,0)); y -= 10*mm

    # Gràfic PV
    if figs.get("pv"):
        pv_h = draw_image_fullwidth(c, figs["pv"], x0, y-60*mm, W - 2*margin)
        y -= (pv_h + 8*mm)
    else:
        y -= 2*mm

    # Canvi de pàgina si cal
    if y < 120*mm:
        c.showPage(); y = H - margin; x0 = margin

    # SOC Li
    if figs.get("soc_li"):
        soc_h = draw_image_fullwidth(c, figs["soc_li"], x0, y-60*mm, W - 2*margin)
        y -= (soc_h + 8*mm)

    # SOC H2
    if figs.get("soc_h2"):
        if y < 120*mm:
            c.showPage(); y = H - margin; x0 = margin
        h2_h = draw_image_fullwidth(c, figs["soc_h2"], x0, y-60*mm, W - 2*margin)
        y -= (h2_h + 8*mm)

    # Fluxos
    if figs.get("flows"):
        if y < 120*mm:
            c.showPage(); y = H - margin; x0 = margin
        flow_h = draw_image_fullwidth(c, figs["flows"], x0, y-60*mm, W - 2*margin)
        y -= (flow_h + 10*mm)

    # Taula costos
    if costs_table_df is not None and not costs_table_df.empty:
        c.showPage(); y = H - margin; x0 = margin
        draw_title(c, "Desglossament de costos per angle", x0, y, size=14); y -= 8*mm
        headers = ["Angle(°)", "Plaques", "Bateries_Li", "Bombones_H2", "Refills", "Capex(€)", "Opex(€)", "Total(€)"]
        col_w = (W - 2*margin) / len(headers)
        c.setFont("Helvetica-Bold", 9); c.setFillColor(colors.black)
        for i, htxt in enumerate(headers):
            c.drawString(x0 + i*col_w + 2, y, htxt)
        y -= 5*mm
        c.setFont("Helvetica", 9)
        for _, r in costs_table_df.iterrows():
            vals = [
                f"{int(r['Angle(°)'])}", f"{int(r['Plaques'])}", f"{int(r['Bateries_Li'])}",
                f"{int(r['Bombones_H2'])}", f"{int(r['Refills'])}",
                f"{r['Capex(€)']:,}", f"{r['Opex(€)']:,}", f"{r['Total(€)']:,}",
            ]
            for i, v in enumerate(vals):
                c.drawString(x0 + i*col_w + 2, y, str(v))
            y -= 5*mm
            if y < 25*mm:
                c.showPage(); y = H - margin; x0 = margin
                c.setFont("Helvetica-Bold", 9)
                for i, htxt in enumerate(headers):
                    c.drawString(x0 + i*col_w + 2, y, htxt)
                y -= 5*mm
                c.setFont("Helvetica", 9)

    c.save()
    pdf_bytes_io.seek(0)
# =======================
# ===== Helpers PDF =====
# =======================
def fig_to_png_bytes(fig, scale=2.0, width=1000, height=500):
    """
    Converteix una figura Plotly a PNG (bytes).
    Si no hi ha 'kaleido' compatible o falla l'export, retorna None.
    """
    try:
        return pio.to_image(fig, format="png", width=width, height=height, scale=scale)
    except Exception as e:
        st.warning(f"No s'han pogut renderitzar els gràfics al PDF (sense Kaleido compatible). El report es generarà sense gràfics.")
        return None

def draw_title(c, text, x, y, size=16, color=(0,0,0)):
    c.setFillColorRGB(*color); c.setFont("Helvetica-Bold", size); c.drawString(x, y, text)

def draw_label(c, text, x, y, size=10, color=(0,0,0)):
    c.setFillColorRGB(*color); c.setFont("Helvetica", size); c.drawString(x, y, text)

def draw_kpi_box(c, x, y, w, h, title, value):
    c.setStrokeColor(colors.HexColor("#E5E7EB"))
    c.setFillColor(colors.white)
    c.roundRect(x, y, w, h, 6, stroke=1, fill=1)
    draw_label(c, title, x+8, y+h-16, size=9, color=(0.3,0.3,0.3))
    draw_title(c, value, x+8, y+h-36, size=14, color=(0,0,0))

def draw_image_fullwidth(c, png_bytes, x, y, max_w):
    img = ImageReader(BytesIO(png_bytes))
    iw, ih = img.getSize()
    scale = max_w / iw
    new_w = max_w
    new_h = ih * scale
    c.drawImage(img, x, y, width=new_w, height=new_h, preserveAspectRatio=True, mask='auto')
    return new_h

def build_pdf_report(
    pdf_bytes_io: BytesIO,
    kpis: dict,
    location_tuple: tuple,
    period_str: str,
    angles_list: list,
    base_load_kw: float,
    fuel_cell_kw_unit: float,
    recommended_angle: int,
    figs: dict,
    costs_table_df: pd.DataFrame
):
    c = canvas.Canvas(pdf_bytes_io, pagesize=A4)
    W, H = A4
    margin = 15*mm
    x0 = margin
    y = H - margin

    # Capçalera
    draw_title(c, "Informe tècnic — Sistema PV + Li + H₂", x0, y, size=18); y -= 10*mm
    lat, lon = location_tuple
    draw_label(c, f"Ubicació: {lat}, {lon}   ·   Període: {period_str}", x0, y); y -= 6*mm
    draw_label(c, f"Angles: {', '.join([str(a)+'°' for a in angles_list])}   ·   Consum base: {base_load_kw} kW   ·   Piles de combustible: {fuel_cell_kw_unit} kW/u", x0, y); y -= 8*mm

    # KPIs
    box_w = (W - 2*margin - 3*6*mm) / 4.0
    box_h = 22*mm
    x = x0
    kpi_items = [
        ("kWh PV total", f"{kpis.get('pv_total_mwh',0):.2f} MWh"),
        ("Export neta", f"{kpis.get('export_mwh',0):.2f} MWh"),
        ("Import post-H₂", f"{kpis.get('import_after_h2_mwh',0):.2f} MWh"),
        ("Refills H₂/any", f"{kpis.get('refills',0)}"),
    ]
    for title, val in kpi_items:
        draw_kpi_box(c, x, y-box_h, box_w, box_h, title, val)
        x += box_w + 6*mm
    y -= (box_h + 10*mm)

    # Recomanació
    draw_label(c, "Recomanació ràpida", x0, y, size=10, color=(0.25,0.25,0.25)); y -= 6*mm
    draw_title(c, f"Angle recomanat: {recommended_angle}° pel millor cost total estimat", x0, y, size=12, color=(0,0,0)); y -= 10*mm

    # Gràfic PV
    if figs.get("pv"):
        pv_h = draw_image_fullwidth(c, figs["pv"], x0, y-60*mm, W - 2*margin)
        y -= (pv_h + 8*mm)

    # Canvi de pàgina si cal
    if y < 120*mm:
        c.showPage(); y = H - margin; x0 = margin

    # SOC Li
    if figs.get("soc_li"):
        soc_h = draw_image_fullwidth(c, figs["soc_li"], x0, y-60*mm, W - 2*margin)
        y -= (soc_h + 8*mm)

    # SOC H2
    if figs.get("soc_h2"):
        if y < 120*mm:
            c.showPage(); y = H - margin; x0 = margin
        h2_h = draw_image_fullwidth(c, figs["soc_h2"], x0, y-60*mm, W - 2*margin)
        y -= (h2_h + 8*mm)

    # Fluxos
    if figs.get("flows"):
        if y < 120*mm:
            c.showPage(); y = H - margin; x0 = margin
        flow_h = draw_image_fullwidth(c, figs["flows"], x0, y-60*mm, W - 2*margin)
        y -= (flow_h + 10*mm)

    # Taula costos
    if costs_table_df is not None and not costs_table_df.empty:
        c.showPage(); y = H - margin; x0 = margin
        draw_title(c, "Desglossament de costos per angle", x0, y, size=14); y -= 8*mm
        headers = ["Angle(°)", "Plaques", "Bateries_Li", "Bombones_H2", "Refills", "Capex(€)", "Opex(€)", "Total(€)"]
        col_w = (W - 2*margin) / len(headers)
        c.setFont("Helvetica-Bold", 9); c.setFillColor(colors.black)
        for i, htxt in enumerate(headers):
            c.drawString(x0 + i*col_w + 2, y, htxt)
        y -= 5*mm
        c.setFont("Helvetica", 9)
        for _, r in costs_table_df.iterrows():
            vals = [
                f"{int(r['Angle(°)'])}", f"{int(r['Plaques'])}", f"{int(r['Bateries_Li'])}",
                f"{int(r['Bombones_H2'])}", f"{int(r['Refills'])}",
                f"{r['Capex(€)']:,}", f"{r['Opex(€)']:,}", f"{r['Total(€)']:,}",
            ]
            for i, v in enumerate(vals):
                c.drawString(x0 + i*col_w + 2, y, str(v))
            y -= 5*mm
            if y < 25*mm:
                c.showPage(); y = H - margin; x0 = margin
                c.setFont("Helvetica-Bold", 9)
                for i, htxt in enumerate(headers):
                    c.drawString(x0 + i*col_w + 2, y, htxt)
                y -= 5*mm
                c.setFont("Helvetica", 9)

    c.save()
    pdf_bytes_io.seek(0)

# =======================
# BOTÓ 3: Excel configuracions + PDF (GENERAR/DESCARREGAR)
# =======================
st.markdown("---")
st.subheader("3) Generar i descarregar")
st.caption("Genera l'Excel de configuracions i l'informe PDF. Després, descarrega'ls amb el botó corresponent.")

# Estat de bytes persistents per als botons de descarrega
if "cfg_excel_bytes" not in st.session_state: st.session_state.cfg_excel_bytes = None
if "report_pdf_bytes" not in st.session_state: st.session_state.report_pdf_bytes = None

# --- Paràmetres de combinacions
li_opts = st.multiselect("Selecciona # de bateries Li (1–24)", options=list(range(1, 25)),
                         default=[2,3,4,5,6,7,8,9,10,11,12], key="cfg_liopts")
h2_opts = st.multiselect("Selecciona # de bombones H₂ (1–11)", options=list(range(1, 12)),
                         default=[4,5], key="cfg_h2opts")

panel_opts_by_angle = {}
if st.session_state.layout_df is not None and not st.session_state.layout_df.empty:
    for _, r in st.session_state.layout_df.iterrows():
        a = int(r["Angle(°)"]); baseN = int(r["N_panells"])
        if baseN <= 0: continue
        panel_opts_by_angle[a] = sorted({max(1, baseN-10), baseN-2, baseN-1, baseN, baseN+1, baseN+2, baseN+10})
else:
    for a in st.session_state.get("pv_angles", [40]):
        panel_opts_by_angle[int(a)] = [20, 40, 60]

angle_panel_selection = {}
for a in st.session_state.get("pv_angles", [40]):
    angle_panel_selection[int(a)] = st.multiselect(
        f"Angle {int(a)}° — # Plaques",
        options=panel_opts_by_angle.get(int(a), [20,40,60]),
        default=panel_opts_by_angle.get(int(a), [20,40,60]),
        key=f"panels_sel_{int(a)}"
    )

# --- Botons GENERA / DESCARREGA en paral·lel
gcol1, gcol2, dcol1, dcol2 = st.columns([1,1,1,1])

with gcol1:
    generate_excel = st.button("🛠️ Genera Excel", use_container_width=True)
with gcol2:
    generate_pdf = st.button("🛠️ Genera PDF", use_container_width=True)

with dcol1:
    st.download_button(
        "⬇️ Descarrega Excel",
        data=st.session_state.cfg_excel_bytes or b"",
        file_name="configuracions_per_angle.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        disabled=(st.session_state.cfg_excel_bytes is None),
        use_container_width=True
    )
with dcol2:
    st.download_button(
        "⬇️ Descarrega informe (PDF)",
        data=st.session_state.report_pdf_bytes or b"",
        file_name="informe_atom_h2.pdf",
        mime="application/pdf",
        disabled=(st.session_state.report_pdf_bytes is None),
        use_container_width=True
    )

# ---------- GENERAR EXCEL ----------
if generate_excel:
    try:
        angles_cfg = st.session_state.get("pv_angles", [])
        if not angles_cfg:
            st.error("Selecciona almenys un angle a la part superior.")
        elif not li_opts or not h2_opts:
            st.error("Tria almenys un valor per a bateries Li i bombones H₂.")
        else:
            # Piles de combustible segons P_max (del form)
            p_max_kw = st.session_state.get("li_pmax", 2.0)
            n_fuel_cells = int(math.ceil(p_max_kw / float(st.session_state.costs["FUEL_CELL_UNIT_KW"]))) if p_max_kw > 0 else 0
            fuel_cells_cost = n_fuel_cells * float(st.session_state.costs["FUEL_CELL_COST"])

            out = BytesIO()
            with pd.ExcelWriter(out, engine="openpyxl") as writer:
                for ang in angles_cfg:
                    raw = fetch_pvgis(lat=st.session_state["pv_lat"], lon=st.session_state["pv_lon"],
                                      start_year=int(st.session_state["pv_sy"]), end_year=int(st.session_state["pv_ey"]),
                                      raddb="PVGIS-SARAH3", pvcalc=1, peak_kw=float(st.session_state["pv_peak"]),
                                      loss_pct=float(st.session_state["pv_loss"]), angle=float(ang), aspect=float(st.session_state["pv_az"]))
                    base = transform_df(raw)

                    rows_cfg = []
                    panels_list = angle_panel_selection.get(int(ang), panel_opts_by_angle.get(int(ang), [20,40,60]))

                    for N_panels in sorted(set(int(x) for x in panels_list)):
                        df_pan = base.copy(); df_pan["P_kW"] = df_pan["P_kW"] * N_panels
                        for n_li in sorted(set(li_opts)):
                            cap_li_kwh_cfg = int(n_li) * float(st.session_state.costs["LI_BATT_UNIT_KWH"])
                            sim_li, dt_hours = simulate_battery(
                                df=df_pan, cap_kwh=cap_li_kwh_cfg,
                                soc_ini=st.session_state["li_soc_ini"], soc_min=st.session_state["li_soc_min"],
                                p_ch_max=st.session_state["li_pch"], p_dis_max=st.session_state["li_pdis"],
                                eta_c=st.session_state["li_eta_c"], eta_d=st.session_state["li_eta_d"],
                                load_kw_fixed=st.session_state["li_load"]
                            )
                            v = resample_for_view(sim_li.copy(), gran="Horària").reset_index()
                            if "soc_%" not in v.columns: v["soc_%"] = st.session_state["li_soc_min"] * 100.0

                            for n_h2 in sorted(set(h2_opts)):
                                cap_h2_kwh_cfg = int(n_h2) * float(st.session_state.costs["H2_CYL_KWH"])
                                vh2 = simulate_h2(
                                    df_hourly=v, dt_hours=dt_hours, cap_h2_kwh=cap_h2_kwh_cfg,
                                    soc_h2_init_pct=st.session_state["h2_soc_ini"],
                                    eta_charge=st.session_state["h2_eta_c"], eta_discharge=st.session_state["h2_eta_d"],
                                    load_kw_fixed=st.session_state["li_load"],
                                    soc_li_min_frac=st.session_state["li_soc_min"],
                                    refill_threshold_frac=(st.session_state["h2_thr"]/100.0)
                                )
                                refills = int(vh2["h2_refill_event"].sum()) if "h2_refill_event" in vh2.columns else 0

                                cost_pan = N_panels * float(st.session_state.costs["PANEL_COST"])
                                cost_li  = int(n_li) * float(st.session_state.costs["LI_BATT_UNIT_COST"])
                                cost_h2c = int(n_h2) * float(st.session_state.costs["H2_CYL_COST"])
                                cost_ref = refills * float(st.session_state.costs["REFILL_COST_PER_CYL"]) * int(n_h2)
                                cost_fc  = fuel_cells_cost

                                transm_press_cost = float(st.session_state.costs["TRANSMISSOR_PRESS"]) * (int(n_h2) / 2.0)
                                valvula_cost = float(st.session_state.costs["VALVULA"]) * (int(n_h2) / 2.0)

                                capex_base = (
                                    float(st.session_state.costs["CONTAINER"]) + float(st.session_state.costs["RACK"]) + cost_pan + cost_li + cost_h2c + cost_fc +
                                    float(st.session_state.costs["MPPT"]) + float(st.session_state.costs["CABLE"]) * N_panels + float(st.session_state.costs["REST_SOLAR"]) + float(st.session_state.costs["SAFETY_RELAY"]) + float(st.session_state.costs["DETECTOR_H2"]) +
                                    float(st.session_state.costs["REST_SAFETY"]) + transm_press_cost + valvula_cost + float(st.session_state.costs["CONTROL_HEATERS"]) +
                                    float(st.session_state.costs["HEATERS"]) * n_h2 + float(st.session_state.costs["PT100"]) * n_h2 + float(st.session_state.costs["ELECTROLYZER"]) + float(st.session_state.costs["CAUDALIMETRE"]) + float(st.session_state.costs["STEP_DOWN"]) +
                                    float(st.session_state.costs["ANTENA"]) + float(st.session_state.costs["PLC"]) + float(st.session_state.costs["STEPDOWN_CONTROL"]) + float(st.session_state.costs["STEPDOWN_GENERACIO"]) +
                                    float(st.session_state.costs["VALVULA_FUELCELL"]) * n_fuel_cells + float(st.session_state.costs["STEPUP"]) * n_fuel_cells +
                                    float(st.session_state.costs["REDUNDANCEMODULE"]) + float(st.session_state.costs["REDUCTOR_PRESS"]) + float(st.session_state.costs["SENSOR_PRESS"]) + float(st.session_state.costs["CLIMATITZACIO"]) + float(st.session_state.costs["OTHERS"])
                                )
                                capex = capex_base / (1 - float(st.session_state.costs["MANUFACTORING_%"])) + float(st.session_state.costs["MONTATGE_COST"])
                                opex  = cost_ref + float(st.session_state.costs["OPEX_BASE"])
                                total = (capex + opex) / (1 - float(st.session_state.costs["MARGE"]))

                                rows_cfg.append({
                                    "Angle(°)": int(ang),
                                    "Plaques": int(N_panels),
                                    "Bateries_Li": int(n_li),
                                    "Li_kWh": round(cap_li_kwh_cfg, 1),
                                    "Bombones_H2": int(n_h2),
                                    "H2_kWh": round(cap_h2_kwh_cfg, 1),
                                    "Refills": int(refills),
                                    "FuelCells": int(n_fuel_cells),
                                    "Cost_plaques(€)": int(cost_pan),
                                    "Cost_Li(€)": int(cost_li),
                                    "Cost_bombones(€)": int(cost_h2c),
                                    "Cost_refills(€)": int(cost_ref),
                                    "Cost_fuelcells(€)": int(cost_fc),
                                    "Capex(€)": round(capex, 2),
                                    "Opex(€)": round(opex, 2),
                                    "Total(€)": round(total, 2),
                                })

                    df_cfg = pd.DataFrame(rows_cfg).sort_values(
                        ["Total(€)", "Plaques", "Bateries_Li", "Bombones_H2"]
                    ).reset_index(drop=True)
                    df_cfg.to_excel(writer, sheet_name=f"{int(ang)}deg", index=False)

            st.session_state.cfg_excel_bytes = out.getvalue()
            st.success("Excel de configuracions generat. Ara pots descarregar-lo amb el botó de la dreta.")

    except Exception as e:
        st.error(f"Error generant l'Excel de configuracions: {e}")

# ---------- GENERAR PDF ----------
if generate_pdf:
    try:
        # Necessitem l'Excel generat per trobar el "millor per angle"
        if not st.session_state.cfg_excel_bytes:
            st.info("Primer genera l'Excel (botó 'Genera Excel') per poder crear l'informe PDF.")
        else:
            # Llegeix totes les pestanyes
            all_rows_concat = []
            with pd.ExcelFile(BytesIO(st.session_state.cfg_excel_bytes)) as xf:
                for sh in xf.sheet_names:
                    df_sh = pd.read_excel(xf, sheet_name=sh)
                    all_rows_concat.append(df_sh)
            cfg_all = pd.concat(all_rows_concat, ignore_index=True)
            best_per_angle = cfg_all.sort_values("Total(€)").groupby("Angle(°)", as_index=False).first()
            best_row = best_per_angle.iloc[0]
            angle_rec = int(best_row["Angle(°)"])
            n_li_rec  = int(best_row["Bateries_Li"])
            n_h2_rec  = int(best_row["Bombones_H2"])
            N_panels_rec = int(best_row["Plaques"])

            # Re-simulació de l’angle recomanat per als KPIs i (si podem) gràfics
            raw = fetch_pvgis(
                lat=st.session_state["pv_lat"], lon=st.session_state["pv_lon"],
                start_year=int(st.session_state["pv_sy"]), end_year=int(st.session_state["pv_ey"]),
                raddb="PVGIS-SARAH3", pvcalc=1, peak_kw=float(st.session_state["pv_peak"]),
                loss_pct=float(st.session_state["pv_loss"]), angle=float(angle_rec), aspect=float(st.session_state["pv_az"])
            )
            base = transform_df(raw)
            base["P_kW"] = base["P_kW"] * N_panels_rec

            cap_li_kwh_cfg = n_li_rec * float(st.session_state.costs["LI_BATT_UNIT_KWH"])
            sim_li, dt_hours = simulate_battery(
                df=base, cap_kwh=cap_li_kwh_cfg,
                soc_ini=st.session_state["li_soc_ini"], soc_min=st.session_state["li_soc_min"],
                p_ch_max=st.session_state["li_pch"], p_dis_max=st.session_state["li_pdis"],
                eta_c=st.session_state["li_eta_c"], eta_d=st.session_state["li_eta_d"],
                load_kw_fixed=st.session_state["li_load"]
            )
            v = resample_for_view(sim_li.copy(), gran="Horària").reset_index()

            cap_h2_kwh_cfg = n_h2_rec * float(st.session_state.costs["H2_CYL_KWH"])
            vh2 = simulate_h2(
                df_hourly=v, dt_hours=dt_hours, cap_h2_kwh=cap_h2_kwh_cfg,
                soc_h2_init_pct=st.session_state["h2_soc_ini"],
                eta_charge=st.session_state["h2_eta_c"], eta_discharge=st.session_state["h2_eta_d"],
                load_kw_fixed=st.session_state["li_load"],
                soc_li_min_frac=st.session_state["li_soc_min"],
                refill_threshold_frac=(st.session_state["h2_thr"]/100.0)
            )

            # KPIs
            pv_total_mwh = base["P_kW"].sum() * dt_hours / 1000.0
            export_mwh = max(0.0, v.get("export_kw", pd.Series(0)).sum() * dt_hours / 1000.0)
            import_after_h2_mwh = max(0.0, vh2.get("import_after_h2_kw", pd.Series(0)).sum() * dt_hours / 1000.0)
            refills = int(vh2["h2_refill_event"].sum()) if "h2_refill_event" in vh2.columns else 0
            kpis = {
                "pv_total_mwh": pv_total_mwh,
                "export_mwh": export_mwh,
                "import_after_h2_mwh": import_after_h2_mwh,
                "refills": refills
            }

            # Figures -> PNG bytes (si possible)
            flows_cols = ["pv_kw","load_kw","import_kw","export_kw","bat_charge_kw","bat_discharge_kw"]
            figs = {}
            try:
                fig_pv = px.line(v, x="datetime", y="pv_kw", title="Generació fotovoltaica (kW) — horària")
                pv_img = fig_to_png_bytes(fig_pv)
                if pv_img: figs["pv"] = pv_img

                fig_soc = px.line(v, x="datetime", y="soc_%", title="SOC bateria Li (%) — horària")
                soc_img = fig_to_png_bytes(fig_soc)
                if soc_img: figs["soc_li"] = soc_img

                vh2_plot = vh2.copy()
                vh2_plot["h2_soc_%"] = np.where(cap_h2_kwh_cfg > 0, 100.0 * vh2_plot["h2_soc_kwh"] / cap_h2_kwh_cfg, 0.0)
                fig_h2 = px.line(vh2_plot, x="datetime", y="h2_soc_%", title="SOC H₂ (%) — horària")
                h2_img = fig_to_png_bytes(fig_h2)
                if h2_img: figs["soc_h2"] = h2_img

                vf = v[["datetime"] + [c for c in flows_cols if c in v.columns]].copy()
                mf = vf.melt(id_vars=["datetime"], var_name="Sèrie", value_name="kW")
                fig_flow = px.line(mf, x="datetime", y="kW", color="Sèrie", title="Fluxos de potència (kW)")
                flow_img = fig_to_png_bytes(fig_flow)
                if flow_img: figs["flows"] = flow_img
            except Exception:
                pass  # ja hem avisat a fig_to_png_bytes si falla

            # Taula resum (millor per angle)
            costs_table_df = best_per_angle[[
                "Angle(°)","Plaques","Bateries_Li","Bombones_H2","Refills","Capex(€)","Opex(€)","Total(€)"
            ]].copy()

            loc_tuple = (st.session_state["pv_lat"], st.session_state["pv_lon"])
            period_str = f"{int(st.session_state['pv_sy'])}–{int(st.session_state['pv_ey'])}"
            angles_list = list(sorted(set(cfg_all["Angle(°)"].astype(int).tolist())))
            base_load_kw = st.session_state["li_load"]
            fuel_cell_kw_unit = st.session_state.costs["FUEL_CELL_UNIT_KW"]

            pdf_buf = BytesIO()
            build_pdf_report(
                pdf_bytes_io=pdf_buf,
                kpis=kpis,
                location_tuple=loc_tuple,
                period_str=period_str,
                angles_list=angles_list,
                base_load_kw=base_load_kw,
                fuel_cell_kw_unit=fuel_cell_kw_unit,
                recommended_angle=angle_rec,
                figs=figs,
                costs_table_df=costs_table_df
            )
            st.session_state.report_pdf_bytes = pdf_buf.getvalue()
            st.success("Informe PDF generat. Ara pots descarregar-lo amb el botó de la dreta.")
    except Exception as e:
        st.error(f"Error generant l'informe PDF: {e}")
