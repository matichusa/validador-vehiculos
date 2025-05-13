# Validador y corrector de Excel para carga de vehículos con estética preservada
import streamlit as st
import pandas as pd
from io import BytesIO
from difflib import get_close_matches
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.worksheet.worksheet import Worksheet

# ---------------- Funciones de validación ----------------
def limpiar_dominio(valor):
    if not isinstance(valor, str):
        return valor
    return valor.replace(" ", "").replace("-", "").replace("/", "").upper()

def validar_aproximado(valor, opciones):
    if valor is None:
        return valor, True, ""
    match = get_close_matches(str(valor).strip().title(), opciones, n=1, cutoff=0.75)
    return (match[0], True, "") if match else (valor, False, "Valor no válido")

def validar_entero(valor):
    try:
        return int(float(valor)), True, ""
    except:
        return valor, False, "Número entero inválido"

def validar_fecha(valor):
    try:
        return pd.to_datetime(valor).strftime("%d/%m/%Y"), True, ""
    except:
        return valor, False, "Fecha inválida"

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
    col_map = {col: idx for idx, col in enumerate(encabezados)}

    errores = []
    corregidos = []
    fill_red = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    font_white = Font(color="FFFFFF", bold=True)

    for row in ws.iter_rows(min_row=9, max_row=ws.max_row, max_col=len(encabezados)):
        for cell in row:
            col_name = encabezados[cell.column - 1]
            val = cell.value
            original_val = val

            if col_name == "Dominio" and isinstance(val, str):
                nuevo = limpiar_dominio(val)
                if nuevo != val:
                    corregidos.append((cell.row, col_name, val, nuevo))
                    cell.value = nuevo

            elif col_name in valores_validos:
                nuevo, ok, motivo = validar_aproximado(val, valores_validos[col_name])
                if not ok:
                    cell.fill = fill_red
                    cell.font = font_white
                    errores.append((cell.row, col_name, val, motivo))
                elif nuevo != val:
                    corregidos.append((cell.row, col_name, val, nuevo))
                    cell.value = nuevo

            elif col_name in columnas_fecha:
                nuevo, ok, motivo = validar_fecha(val)
                if not ok:
                    cell.fill = fill_red
                    cell.font = font_white
                    errores.append((cell.row, col_name, val, motivo))
                elif nuevo != val:
                    corregidos.append((cell.row, col_name, val, nuevo))
                    cell.value = nuevo

            elif col_name == "Odómetro":
                nuevo, ok, motivo = validar_entero(val)
                if not ok:
                    cell.fill = fill_red
                    cell.font = font_white
                    errores.append((cell.row, col_name, val, motivo))
                elif nuevo != val:
                    corregidos.append((cell.row, col_name, val, nuevo))
                    cell.value = nuevo

    # Mostrar resumen de errores y correcciones
    if errores:
        st.markdown("### ❌ Valores no corregidos")
        st.dataframe(pd.DataFrame(errores, columns=["Fila", "Columna", "Valor original", "Observación"]))

    if corregidos:
        st.markdown("### ✅ Valores corregidos automáticamente")
        st.dataframe(pd.DataFrame(corregidos, columns=["Fila", "Columna", "Valor original", "Valor corregido"]))

    # Agregar hoja log de errores
    if errores:
        log_ws = wb.create_sheet("Log de Errores")
        log_ws.append(["Fila", "Columna", "Valor original", "Observación"])
        for fila, columna, valor, observacion in errores:
            log_ws.append([fila, columna, valor, observacion])

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.download_button(
        label="📥 Descargar archivo validado",
        data=output,
        file_name="vehiculos_validado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.success("Archivo validado conservando su estética. Se agregó una hoja con el log de errores si se detectaron.")


