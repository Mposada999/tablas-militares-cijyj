import math
import os
from dataclasses import dataclass
from datetime import datetime

import streamlit as st
from openpyxl import Workbook, load_workbook


# =============================
#  EXCEL (BASE DE DATOS)
# =============================
EXCEL_PATH = "tablas_militares_cijyj_registro.xlsx"
SHEET_NAME = "Inspecciones"

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

# AQL informativo (solo pantalla)
AC_CRIT = 0
RE_CRIT = 1
PCT_MAY_AC = 0.03
PCT_MAY_RE = 0.05
PCT_MEN_AC = 0.06
PCT_MEN_RE = 0.10

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

def build_plan(presentacion: str, lote: int, nivel: str) -> PlanMuestreo:
    codigo, n = get_sample_size(lote, nivel)
    ac_may, re_may = ac_re_from_pct(n, PCT_MAY_AC, PCT_MAY_RE)
    ac_men, re_men = ac_re_from_pct(n, PCT_MEN_AC, PCT_MEN_RE)
    return PlanMuestreo(
        presentacion=presentacion,
        lote=lote,
        nivel=nivel,
        codigo=codigo,
        n=n,
        ac_crit=AC_CRIT, re_crit=RE_CRIT,
        ac_may=ac_may, re_may=re_may,
        ac_men=ac_men, re_men=re_men
    )

def plan_or_empty(presentacion: str, qty: int, nivel: str):
    if qty <= 0:
        return None
    return build_plan(presentacion, qty, nivel)

def fill(plan, prefix):
    if plan is None:
        return {f"Cant_{prefix}": 0, f"Codigo_{prefix}": "", f"n_{prefix}": ""}
    return {f"Cant_{prefix}": plan.lote, f"Codigo_{prefix}": plan.codigo, f"n_{prefix}": plan.n}


# =============================
#  STREAMLIT APP
# =============================
st.set_page_config(page_title="Tablas militares | CI JYJ", layout="centered")
st.title("🧪 Tablas militares (MIL-STD-105E) | CI JYJ / Labsens")

st.markdown("Diligencia la recepción, calcula el plan de muestreo y guarda **1 registro por corrida**.")

with st.form("form_recepcion"):
    st.subheader("1) Datos generales")
    operario = st.text_input("Operario (nombre)")
    proveedor = st.text_input("Proveedor")
    fragancia = st.text_input("Fragancia")
    num_lote_proveedor = st.text_input("Número de lote proveedor")
    libras_caneca = st.number_input("Libras de la caneca (si no aplica, 0)", min_value=0.0, step=0.5)

    nivel = st.selectbox("Nivel de inspección", ["I", "II", "III"], index=1)

    st.subheader("2) Cantidades recibidas")
    q500 = st.number_input("Cantidad recibida 500 g", min_value=0, step=1)
    q220 = st.number_input("Cantidad recibida 220 g", min_value=0, step=1)
    q30  = st.number_input("Cantidad recibida 30 g",  min_value=0, step=1)

    st.subheader("3) Registro final")
    calidad_esperada = st.selectbox("¿La fragancia cumplió con la calidad esperada?", ["Si", "No"])
    observaciones = st.text_area("Observaciones")

    submitted = st.form_submit_button("✅ Calcular y Guardar")

if submitted:
    if (q500 + q220 + q30) == 0:
        st.error("Debes ingresar al menos una cantidad > 0.")
        st.stop()

    p500 = plan_or_empty("500 g", q500, nivel)
    p220 = plan_or_empty("220 g", q220, nivel)
    p30  = plan_or_empty("30 g",  q30,  nivel)

    st.success("Plan calculado ✅")

    st.subheader("📌 Plan de muestreo por presentación (informativo)")
    for plan in [p500, p220, p30]:
        if plan is None:
            continue
        st.markdown(
            f"""
**{plan.presentacion}** | Lote={plan.lote} | Nivel={plan.nivel}  
- Letra código: **{plan.codigo}**  
- Tamaño de muestra (n): **{plan.n}**  
- Crítico: Ac={plan.ac_crit} Re={plan.re_crit} (cero tolerancia)  
- Mayor: Ac={plan.ac_may} Re={plan.re_may}  
- Menor: Ac={plan.ac_men} Re={plan.re_men}  
            """
        )

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
    row.update(fill(p500, "500g"))
    row.update(fill(p220, "220g"))
    row.update(fill(p30,  "30g"))

    append_row(row)
    st.success(f"✅ Guardado en Excel: {EXCEL_PATH}")

    # Descargar el Excel desde la app
    if os.path.exists(EXCEL_PATH):
        with open(EXCEL_PATH, "rb") as f:
            st.download_button(
                "⬇️ Descargar Excel",
                data=f,
                file_name=EXCEL_PATH,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
