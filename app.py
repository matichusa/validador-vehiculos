
# validador_excel_vehiculos.py
import streamlit as st
import pandas as pd
from io import BytesIO
from difflib import get_close_matches
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from datetime import datetime

# ---------------- Funciones de validación ----------------
def limpiar_dominio(valor):
    if not isinstance(valor, str):
        return valor
    return ''.join(filter(str.isalnum, valor.upper()))

def normalizar(valor):
    if valor is None:
        return ""
    return str(valor).strip().lower().replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")

def validar_aproximado(valor, opciones):
    val_norm = normalizar(valor)
    opciones_norm = [normalizar(o) for o in opciones]
    if val_norm in opciones_norm:
        idx = opciones_norm.index(val_norm)
        return opciones[idx], True, ""
    coincidencias = get_close_matches(val_norm, opciones_norm, n=1, cutoff=0.8)
    if coincidencias:
        idx = opciones_norm.index(coincidencias[0])
        return opciones[idx], True, ""
    return valor, False, "Valor no válido"

def validar_entero(valor):
    if pd.isna(valor) or valor == "":
        return valor, True, ""
    try:
        return int(float(valor)), True, ""
    except:
        return valor, False, "Número entero inválido"

def validar_decimal(valor):
    if pd.isna(valor) or valor == "":
        return valor, True, ""
    try:
        val = round(float(valor), 1)
        return val, True, ""
    except:
        return valor, False, "Número decimal inválido"

def validar_fecha(valor):
    if pd.isna(valor) or valor == "":
        return "", True, ""
    try:
        if isinstance(valor, datetime):
            return valor.strftime("%d/%m/%Y"), True, ""
        elif isinstance(valor, str):
            valor = valor.strip().lower().replace(" del ", " ").replace(" de ", " ")
            meses = {
                "enero": "01", "ene": "01", "january": "01", "jan": "01",
                "febrero": "02", "feb": "02", "february": "02",
                "marzo": "03", "mar": "03", "march": "03",
                "abril": "04", "apr": "04", "april": "04",
                "mayo": "05", "may": "05",
                "junio": "06", "jun": "06", "june": "06",
                "julio": "07", "jul": "07", "july": "07",
                "agosto": "08", "aug": "08", "august": "08",
                "septiembre": "09", "sep": "09", "sept": "09", "september": "09",
                "octubre": "10", "oct": "10", "october": "10",
                "noviembre": "11", "nov": "11", "november": "11",
                "diciembre": "12", "dic": "12", "dec": "12", "december": "12"
            }
            for mes, num in meses.items():
                if mes in valor:
                    valor = valor.replace(mes, num)
            valor = valor.replace("-", "/").replace(".", "/")
            fecha = pd.to_datetime(valor, dayfirst=True, errors='coerce')
            if pd.isna(fecha):
                raise ValueError("No se pudo convertir")
            return fecha.strftime("%d/%m/%Y"), True, ""
        fecha = pd.to_datetime(valor, dayfirst=True, errors='raise')
        return fecha.strftime("%d/%m/%Y"), True, ""
    except Exception as e:
        return valor, False, f"Fecha inválida: {e}"

# ---------------- App Streamlit ----------------
st.title("Validador Excel Vehículos")

file = st.file_uploader("Subí el archivo Excel", type=["xlsx"])

if file:
    wb = load_workbook(file)
    ws = wb.active

    encabezados = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    encabezados_norm = [normalizar(e) for e in encabezados]

    errores = []
    corregidos = []
    cambios_por_columna = {}
    valores_unicos = {"dominio": set()}
    fill_red = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    font_white = Font(color="FFFFFF", bold=True)

    columnas_fecha = [
        "vto póliza", "vto cédula", "vto vtv", "vto ruta",
        "vto gnc", "vto cilindro gnc", "vto senasa"
    ]

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            col_idx = cell.column - 1
            col_name = encabezados_norm[col_idx]
            original_val = cell.value

            if col_name == "dominio":
                nuevo = limpiar_dominio(original_val)
                if nuevo in valores_unicos["dominio"]:
                    cell.fill = fill_red
                    cell.font = font_white
                    errores.append((cell.row, encabezados[col_idx], original_val, "Dominio duplicado"))
                else:
                    valores_unicos["dominio"].add(nuevo)
                    if nuevo != original_val:
                        corregidos.append((cell.row, encabezados[col_idx], original_val, nuevo))
                        cambios_por_columna[encabezados[col_idx]] = cambios_por_columna.get(encabezados[col_idx], 0) + 1
                        cell.value = nuevo

            elif col_name in columnas_fecha:
                nuevo, ok, msg = validar_fecha(original_val)
                if not ok:
                    cell.fill = fill_red
                    cell.font = font_white
                    errores.append((cell.row, encabezados[col_idx], original_val, msg))
                elif nuevo != original_val:
                    corregidos.append((cell.row, encabezados[col_idx], original_val, nuevo))
                    cambios_por_columna[encabezados[col_idx]] = cambios_por_columna.get(encabezados[col_idx], 0) + 1
                    cell.value = nuevo

    st.markdown(f"### 📊 Resumen general")
    st.info(f"Se detectaron **{len(errores)} errores** y se corrigieron automáticamente **{len(corregidos)} valores**.")

    with st.expander("📋 Ver detalles de cambios"):
        if errores:
            st.markdown("### ❌ No corregidos")
            st.dataframe(pd.DataFrame(errores, columns=["Fila", "Columna", "Valor original", "Observación"]))
        if corregidos:
            st.markdown("### ✅ Corregidos")
            st.dataframe(pd.DataFrame(corregidos, columns=["Fila", "Columna", "Valor original", "Valor corregido"]))

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.download_button(
        "📥 Descargar Excel Validado",
        data=output,
        file_name="vehiculos_validado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
