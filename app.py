# Validador y corrector de Excel para carga de vehículos con estética preservada
import streamlit as st
import pandas as pd
from io import BytesIO
from difflib import get_close_matches
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# ---------------- Funciones de validación ----------------
def limpiar_dominio(valor):
    if not isinstance(valor, str):
        return valor
    return valor.replace(" ", "").replace("-", "").replace("/", "").upper()

def validar_aproximado(valor, opciones):
    if valor is None:
        return valor, True
    match = get_close_matches(str(valor).strip().title(), opciones, n=1, cutoff=0.75)
    return (match[0], True) if match else (valor, False)

def validar_entero(valor):
    try:
        return int(float(valor)), True
    except:
        return valor, False

def validar_fecha(valor):
    try:
        return pd.to_datetime(valor).strftime("%d/%m/%Y"), True
    except:
        return valor, False

# ---------------- Configuración ----------------
valores_validos = {
    "Combustible": ["Nafta", "Diésel", "Gas", "Eléctrico"],
    "Med. Uso": ["Kilometros", "Millas", "Horas"],
    "Estado": ["Asignado", "Disponible", "En Taller", "Fuera de Servicio"],
    "Tipo de Cobertura": ["Terceros Completo Estandard", "Tercero Completo Premium", "Todo Riesgo"],
    "Titularidad": ["Propio", "Alquilado", "Leasing", "Prendario"]
}
columnas_fecha = ["Vto Póliza", "Vto Cédula", "Vto VTV", "Vto Ruta", "Vto GNC", "Vto Cilindro GNC", "Vto Senasa"]

# ---------------- App Streamlit ----------------
st.title("Validador de Archivo Excel - Estética Original")
file = st.file_uploader("Subí el archivo Excel original", type=[".xlsx"])

if file:
    wb = load_workbook(file)
    ws = wb[wb.sheetnames[0]]

    encabezados = [cell.value for cell in ws[7]]  # Fila 8 (indexada desde 1)
    col_map = {col: idx + 1 for idx, col in enumerate(encabezados)}

    errores = []
    fill_red = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    font_white = Font(color="FFFFFF", bold=True)

    for row in ws.iter_rows(min_row=9, max_row=ws.max_row, max_col=len(encabezados)):
        for col_name, col_idx in col_map.items():
            cell = ws.cell(row=row[0].row, column=col_idx)
            val = cell.value

            if col_name == "Dominio" and isinstance(val, str):
                nuevo = limpiar_dominio(val)
                if nuevo != val:
                    cell.value = nuevo

            elif col_name in valores_validos:
                nuevo, ok = validar_aproximado(val, valores_validos[col_name])
                if not ok:
                    cell.fill = fill_red
                    cell.font = font_white

            elif col_name in columnas_fecha:
                _, ok = validar_fecha(val)
                if not ok:
                    cell.fill = fill_red
                    cell.font = font_white

            elif col_name == "Odómetro":
                _, ok = validar_entero(val)
                if not ok:
                    cell.fill = fill_red
                    cell.font = font_white

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.download_button(
        label="📥 Descargar archivo validado",
        data=output,
        file_name="vehiculos_validado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.success("Archivo validado conservando su estética.")

