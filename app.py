import math
import os
from dataclasses import dataclass
from datetime import datetime

import streamlit as st
import pandas as pd
from openpyxl import Workbook, load_workbook


# =============================
#  EXCEL (BASE DE DATOS)
# =============================
EXCEL_PATH = "tablas_militares_cijyj_registro.xlsx"
SHEET_NAME = "Inspecciones"

# 1 fila por corrida
HEADERS = [
    "FechaHora",
    "Operario",
    "Proveedor",
    "Fragancia",
    "NumeroLoteProveedor",
    "LibrasCaneca",
    "NivelInspeccion",

    "Cant_500g", "Codigo_500g", "n_500g",
    "Cant_220g", "Codigo_220g", "n_220g",
    "Cant_30g",  "Codigo_30g",  "n_30g",

    "CalidadEsperada",
    "Observaciones"
]

def ensure_workbook(path=EXCEL_PATH):
    if os.path.exists(path):
        wb = load_workbook(path)
        if SHEET_NAME not in wb.sheetnames:
            ws = wb.create_sheet(SHEET_NAME)
            ws.append(HEADERS)
            wb.save(path)
        else:
            ws = wb[SHEET_NAME]
            if ws.max_row == 0:
                ws.append(HEADERS)
                wb.save(path)
        return

    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    ws.append(HEADERS)
    wb.save(path)

def append_row(row_dict: dict, path=EXCEL_PATH):
    ensure_workbook(path)
    wb = load_workbook(path)
    ws = wb[SHEET_NAME]
    row = [row_dict.get(h, "") for h in HEADERS]
    ws.append(row)
    wb.save(path)


# -----------------------------
# MIL-STD-105E - TABLE I (letra + n)
# -----------------------------
LOT_RANGES = [
    (2, 8,     {"I": "A", "II": "A", "III": "B"}),
    (9, 15,    {"I": "A", "II": "B", "III": "C"}),
    (16, 25,   {"I": "B", "II": "C", "III": "D"}),
    (26, 50,   {"I": "C", "II": "D", "III": "E"}),

    (51, 90,   {"I": "D", "II": "F", "III": "G"}),
    (91, 150,  {"I": "D", "II": "F", "III": "G"}),
    (151, 280, {"I": "E", "II": "G", "III": "H"}),
    (281, 500, {"I": "F", "II": "H", "III": "J"}),
    (501, 1200, {"I": "G", "II": "J", "III": "K"}),
    (1201, 3200, {"I": "H", "II": "K", "III": "L"}),
    (3201, 10000, {"I": "J", "II": "L", "III": "M"}),
    (10001, 35000, {"I": "K", "II": "M", "III": "N"}),
    (35001, 150000, {"I": "L", "II": "N", "III": "P"}),
    (150001, 500000, {"I": "M", "II": "P", "III": "Q"}),
    (500001, 10**18, {"I": "N", "II": "Q", "III": "R"}),
]

SAMPLE_SIZE_BY_CODE = {
    "A": 2, "B": 3, "C": 5, "D": 8, "E": 13, "F": 20, "G": 32, "H": 50,
    "J": 80, "K": 125, "L": 200, "M": 315, "N": 500, "P": 800, "Q": 1250, "R": 2000
}

# -----------------------------
# AQL (informativo)
# OJO: estos porcentajes son "política Labsens" inicial.
# Se pueden ajustar en el sidebar sin tocar código.
# -----------------------------
DEFAULT_AC_CRIT = 0
DEFAULT_RE_CRIT = 1
DEFAULT_PCT_MAY_AC = 0.03
DEFAULT_PCT_MAY_RE = 0.05
DEFAULT_PCT_MEN_AC = 0.06
DEFAULT_PCT_MEN_RE = 0.10


@dataclass
class PlanMuestreo:
    presentacion: str
    lote: int
    nivel: str
    codigo: str
    n: int
    ac_crit: int
    re_crit: int
    ac_may: int
    re_may: int
    ac_men: int
    re_men: int


def get_code_letter(lot_size: int, level: str) -> str:
    for lo, hi, mapping in LOT_RANGES:
        if lo <= lot_size <= hi:
            return mapping[level]
    return "N"

def get_sample_size(lot_size: int, level: str):
    code = get_code_letter(lot_size, level)
    n = SAMPLE_SIZE_BY_CODE[code]
    return code, min(n, lot_size)

def ac_re_from_pct(n: int, pct_ac: float, pct_re: float):
    ac = math.floor(pct_ac * n)
    re = math.ceil(pct_re * n)
    if re <= ac:
        re = ac + 1
    return ac, re

def build_plan(presentacion: str, lote: int, nivel: str,
               ac_crit: int, re_crit: int,
               pct_may_ac: float, pct_may_re: float,
               pct_men_ac: float, pct_men_re: float) -> PlanMuestreo:

    codigo, n = get_sample_size(lote, nivel)
    ac_may, re_may = ac_re_from_pct(n, pct_may_ac, pct_may_re)
    ac_men, re_men = ac_re_from_pct(n, pct_men_ac, pct_men_re)

    return PlanMuestreo(
        presentacion=presentacion,
        lote=lote,
        nivel=nivel,
        codigo=codigo,
        n=n,
        ac_crit=ac_crit, re_crit=re_crit,
        ac_may=ac_may, re_may=re_may,
        ac_men=ac_men, re_men=re_men
    )

def plan_or_none(presentacion: str, qty: int, nivel: str, cfg: dict):
    if qty <= 0:
        return None
    return build_plan(
        presentacion, qty, nivel,
        cfg["ac_crit"], cfg["re_crit"],
        cfg["pct_may_ac"], cfg["pct_may_re"],
        cfg["pct_men_ac"], cfg["pct_men_re"]
    )

def row_fields_from_plan(plan, prefix):
    if plan is None:
        return {f"Cant_{prefix}": 0, f"Codigo_{prefix}": "", f"n_{prefix}": ""}
    return {f"Cant_{prefix}": plan.lote, f"Codigo_{prefix}": plan.codigo, f"n_{prefix}": plan.n}

def plans_to_table(p500, p220, p30):
    rows = []
    for label, plan in [("500 g", p500), ("220 g", p220), ("30 g", p30)]:
        if plan is None:
            rows.append({
                "Presentación": label,
                "Lote (und)": 0,
                "Código": "—",
                "Muestra (n)": "—",
                "Crítico (Ac/Re)": "—",
                "Mayor (Ac/Re)": "—",
                "Menor (Ac/Re)": "—",
            })
        else:
            rows.append({
                "Presentación": label,
                "Lote (und)": plan.lote,
                "Código": plan.codigo,
                "Muestra (n)": plan.n,
                "Crítico (Ac/Re)": f"{plan.ac_crit}/{plan.re_crit}",
                "Mayor (Ac/Re)": f"{plan.ac_may}/{plan.re_may}",
                "Menor (Ac/Re)": f"{plan.ac_men}/{plan.re_men}",
            })
    return pd.DataFrame(rows)


# =============================
#  STREAMLIT UI
# =============================
st.set_page_config(page_title="Tablas militares | CI JYJ / Labsens", layout="centered")

st.title("🧪 Tablas militares (MIL-STD-105E) | CI JYJ / Labsens")
st.caption("Diligencia la recepción, calcula el plan de muestreo y guarda **1 registro por corrida**.")

# Sidebar: configuración de AQL (informativo)
with st.sidebar:
    st.subheader("⚙️ Parámetros AQL (informativo)")
    st.caption("Estos valores son política interna (puedes ajustarlos).")
    ac_crit = st.number_input("Crítico - Ac", min_value=0, value=DEFAULT_AC_CRIT, step=1)
    re_crit = st.number_input("Crítico - Re", min_value=1, value=DEFAULT_RE_CRIT, step=1)

    st.markdown("**Mayor (porcentaje sobre n)**")
    pct_may_ac = st.number_input("Mayor - % para ACEPTAR", min_value=0.0, max_value=1.0, value=DEFAULT_PCT_MAY_AC, step=0.01, format="%.2f")
    pct_may_re = st.number_input("Mayor - % para RECHAZAR", min_value=0.0, max_value=1.0, value=DEFAULT_PCT_MAY_RE, step=0.01, format="%.2f")

    st.markdown("**Menor (porcentaje sobre n)**")
    pct_men_ac = st.number_input("Menor - % para ACEPTAR", min_value=0.0, max_value=1.0, value=DEFAULT_PCT_MEN_AC, step=0.01, format="%.2f")
    pct_men_re = st.number_input("Menor - % para RECHAZAR", min_value=0.0, max_value=1.0, value=DEFAULT_PCT_MEN_RE, step=0.01, format="%.2f")

cfg = {
    "ac_crit": int(ac_crit),
    "re_crit": int(re_crit),
    "pct_may_ac": float(pct_may_ac),
    "pct_may_re": float(pct_may_re),
    "pct_men_ac": float(pct_men_ac),
    "pct_men_re": float(pct_men_re),
}

# Guardamos el plan calculado en session_state para que el usuario lo vea antes de guardar
if "last_plan_df" not in st.session_state:
    st.session_state["last_plan_df"] = None
if "last_plans_obj" not in st.session_state:
    st.session_state["last_plans_obj"] = (None, None, None)

with st.form("form_recepcion"):
    st.subheader("1) Datos generales")
    operario = st.text_input("Operario (nombre)", placeholder="Ej: Juan Carlos")
    proveedor = st.text_input("Proveedor", placeholder="Ej: Superpack")
    fragancia = st.text_input("Fragancia", placeholder="Ej: Fragancia Rosa")
    num_lote_proveedor = st.text_input("Número de lote proveedor", placeholder="Ej: 431278")
    libras_caneca = st.number_input("Libras de la caneca (si no aplica, 0)", min_value=0.0, step=0.5)

    nivel = st.selectbox("Nivel de inspección", ["I", "II", "III"], index=1)

    st.subheader("2) Cantidades recibidas")
    c1, c2, c3 = st.columns(3)
    with c1:
        q500 = st.number_input("500 g", min_value=0, step=1)
    with c2:
        q220 = st.number_input("220 g", min_value=0, step=1)
    with c3:
        q30  = st.number_input("30 g", min_value=0, step=1)

    st.subheader("3) Registro final")
    calidad_esperada = st.selectbox("¿La fragancia cumplió con la calidad esperada?", ["Si", "No"])
    observaciones = st.text_area("Observaciones", placeholder="Ej: N/A o detalle de hallazgos")

    colA, colB = st.columns(2)
    with colA:
        btn_calcular = st.form_submit_button("📌 Calcular plan")
    with colB:
        btn_guardar = st.form_submit_button("💾 Guardar registro")

def calcular_planes(q500, q220, q30, nivel, cfg):
    p500 = plan_or_none("500 g", int(q500), nivel, cfg)
    p220 = plan_or_none("220 g", int(q220), nivel, cfg)
    p30  = plan_or_none("30 g",  int(q30),  nivel, cfg)
    df = plans_to_table(p500, p220, p30)
    return p500, p220, p30, df

def mostrar_resultados(df):
    st.subheader("📌 Plan de muestreo (código, muestra n y AQL informativo)")
    st.dataframe(df, use_container_width=True, hide_index=True)
    st.caption("Crítico/Major/Menor se muestran como Ac/Re (aceptación / rechazo).")

if btn_calcular or btn_guardar:
    if (q500 + q220 + q30) == 0:
        st.error("Debes ingresar al menos una cantidad > 0.")
    else:
        p500, p220, p30, df = calcular_planes(q500, q220, q30, nivel, cfg)
        st.session_state["last_plan_df"] = df
        st.session_state["last_plans_obj"] = (p500, p220, p30)

        # Siempre mostramos el plan cuando el usuario le da a cualquier botón
        mostrar_resultados(df)

        if btn_guardar:
            # Validación mínima (para no guardar basura)
            if not operario.strip() or not proveedor.strip() or not fragancia.strip():
                st.warning("Antes de guardar, completa Operario, Proveedor y Fragancia.")
            else:
                now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                row = {
                    "FechaHora": now,
                    "Operario": operario.strip(),
                    "Proveedor": proveedor.strip(),
                    "Fragancia": fragancia.strip(),
                    "NumeroLoteProveedor": num_lote_proveedor.strip(),
                    "LibrasCaneca": libras_caneca if libras_caneca != 0 else "",
                    "NivelInspeccion": nivel,
                    "CalidadEsperada": calidad_esperada,
                    "Observaciones": observaciones.strip() if observaciones else ""
                }

                row.update(row_fields_from_plan(p500, "500g"))
                row.update(row_fields_from_plan(p220, "220g"))
                row.update(row_fields_from_plan(p30,  "30g"))

                append_row(row)

                st.success(f"✅ Registro guardado (1 fila). Archivo: {EXCEL_PATH}")

                # Descargar el Excel actualizado
                if os.path.exists(EXCEL_PATH):
                    with open(EXCEL_PATH, "rb") as f:
                        st.download_button(
                            "⬇️ Descargar Excel actualizado",
                            data=f,
                            file_name=EXCEL_PATH,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

# Si ya calcularon antes, mostramos el último plan aunque no hayan tocado botones (mejor UX)
if st.session_state["last_plan_df"] is not None and not (btn_calcular or btn_guardar):
    mostrar_resultados(st.session_state["last_plan_df"])
