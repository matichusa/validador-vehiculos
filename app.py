# Validador y corrector de Excel para carga de vehículos
# Aplicación en Streamlit para facilitar su uso

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from difflib import get_close_matches
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

# ---------------- Funciones de validación ----------------
def limpiar_dominio(valor):
    if pd.isna(valor):
        return ""
    return str(valor).replace(" ", "").replace("-", "").replace("/", "").upper()

def validar_aproximado(valor, opciones):
    if pd.isna(valor):
        return np.nan
    valor_str = str(valor).strip().title()
    match = get_close_matches(valor_str, opciones, n=1, cutoff=0.75)
    return match[0] if match else "NO CORREGIDO"

def validar_entero(valor):
    try:
        return int(float(valor)) if not pd.isna(valor) else np.nan
    except:
        return "NO CORREGIDO"

def validar_fecha(valor):
    if pd.isna(valor):
        return np.nan
    try:
        return pd.to_datetime(valor).strftime("%d/%m/%Y")
    except:
        return "NO CORREGIDO"

def convertir_a_titulo(valor):
    if pd.isna(valor):
        return valor
    return str(valor).strip().title()

# ---------------- App Streamlit ----------------
st.title("Validador de Archivo Excel - Vehículos")
st.write("Subí tu archivo de Excel para validar y corregir los datos automáticamente.")

file = st.file_uploader("Cargá el archivo Excel", type=[".xlsx"])

if file:
    df = pd.read_excel(file, header=None)
    header_row = df[df.apply(lambda row: row.astype(str).str.contains("Dominio").any(), axis=1)].index[0]
    df = pd.read_excel(file, header=header_row)

    df_original = df.copy()  # Para comparación

    # Correcciones manuales
    correcciones_manual = {
        "Titularidad": {"Mio": "Propio"},
        "Med. Uso": {"Kilometro": "Kilometros", "Kilómetro": "Kilometros", "km": "Kilometros"},
        "Color": {"Amarillllo": "Amarillo", "amarillo": "Amarillo"}
    }

    opciones_validas = {
        "Combustible": ["Nafta", "Diésel", "Gas", "Eléctrico"],
        "Med. Uso": ["Kilometros", "Millas", "Horas"],
        "Estado": ["Asignado", "Disponible", "En Taller", "Fuera de Servicio"],
        "Tipo Cobertura": ["Terceros Completo Estandard", "Tercero Completo Premium", "Todo Riesgo"],
        "Titularidad": ["Propio", "Alquilado", "Leasing", "Prendario"]
    }

    fechas = ["Vto Póliza", "Vto Cédula", "Vto VTV", "Vto Ruta", "Vto GNC", "Vto Cilindro GNC", "Vto Senasa"]
    columnas_sin_nan = ["Color", "Nro Chasis", "Nro Motor"]

    for col, reemplazos in correcciones_manual.items():
        if col in df.columns:
            df[col] = df[col].replace(reemplazos)

    if "Dominio" in df.columns:
        df["Dominio"] = df["Dominio"].apply(limpiar_dominio)

    if "Codigo - Interno" in df.columns:
        df["Codigo - Interno"] = df["Codigo - Interno"].astype(str).str.strip().str.upper()

    for columna, opciones in opciones_validas.items():
        if columna in df.columns:
            df[columna] = df[columna].apply(lambda x: validar_aproximado(x, opciones))

    for col in df.columns:
        if col not in fechas and col not in {"Codigo - Interno", "Dominio"} and df[col].dtype == object:
            df[col] = df[col].apply(convertir_a_titulo)

    for col in fechas:
        if col in df.columns:
            df[col] = df[col].apply(validar_fecha)

    if "Odómetro" in df.columns:
        df["Odómetro"] = df["Odómetro"].apply(validar_entero)

    for col in columnas_sin_nan:
        if col in df.columns:
            df[col] = df[col].apply(lambda x: "" if pd.isna(x) else x)

    # Comparar cambios
    cambios = (df != df_original) & ~(df.isna() & df_original.isna())
    total_cambios = cambios.sum().sum()
    columnas_con_cambios = cambios.sum()
    resumen_cambios = df_original[cambios].copy()
    resumen_cambios_nuevo = df[cambios].copy()
    resumen = pd.DataFrame({
        'Valor original': resumen_cambios.stack(),
        'Valor corregido': resumen_cambios_nuevo.stack()
    })

    errores_no_corregidos = resumen[resumen['Valor corregido'] == "NO CORREGIDO"]
    resumen_ok = resumen[resumen['Valor corregido'] != "NO CORREGIDO"]

    with st.expander("📋 Ver resumen de cambios detectados"):
        if not errores_no_corregidos.empty:
            st.markdown("### ⚠️ Cambios no corregidos")
            st.dataframe(errores_no_corregidos)

        st.write(f"🔧 Total de celdas corregidas: **{total_cambios}**")
        st.dataframe(columnas_con_cambios[columnas_con_cambios > 0])
        if not resumen_ok.empty:
            st.markdown("### ✅ Cambios realizados")
            st.dataframe(resumen_ok)

            # Agregar botón para descargar el log de cambios
            output_log = BytesIO()
            resumen.to_excel(output_log)
            output_log.seek(0)
            st.download_button(
                label="Descargar log de cambios",
                data=output_log,
                file_name="log_cambios.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # Descargar archivo corregido con formato y hoja adicional de log
    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)
    wb = load_workbook(filename=output)
    ws = wb.active

    # Marcar errores en hoja principal
    for (idx, row), (colname, coldata) in zip(enumerate(df.iterrows(), start=2), df.items()):
        for col_index, val in enumerate(row[1], start=1):
            if val == "NO CORREGIDO":
                cell = ws.cell(row=idx, column=col_index)
                cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                cell.font = Font(color="FFFFFF", bold=True)

    # Agregar hoja con el resumen de cambios
    ws_log = wb.create_sheet(title="Log de Cambios")
    for r in dataframe_to_rows(resumen.reset_index(), index=False, header=True):
        ws_log.append(r)

    corrected_output = BytesIO()
    wb.save(corrected_output)
    corrected_output.seek(0)

    st.download_button(
        label="Descargar archivo corregido",
        data=corrected_output,
        file_name="vehiculos_corregido.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.success("Archivo procesado correctamente.")

