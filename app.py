# Validador y corrector de Excel para carga de vehículos con estética preservada
import streamlit as st
import pandas as pd
from io import BytesIO
from difflib import get_close_matches
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# ---------------- Funciones de validación ----------------
import unicodedata

def normalizar_columna(nombre):
    if not isinstance(nombre, str):
        return ""
    nombre = unicodedata.normalize("NFKD", nombre).encode("ascii", "ignore").decode("utf-8")
    return nombre.strip().lower()
def limpiar_dominio(valor):
    if not isinstance(valor, str):
        return valor
    return ''.join(c for c in valor.upper() if c.isalnum())

def normalizar(valor):
    if valor is None:
        return ""
    return str(valor).strip().lower().replace("á", "a").replace("é", "e")        .replace("í", "i").replace("ó", "o").replace("ú", "u")

def titulo_propio(valor):
    if not isinstance(valor, str):
        return valor
    return ' '.join(p.capitalize() for p in valor.strip().split())

def mayusculas(valor):
    if not isinstance(valor, str):
        return valor
    return valor.strip().upper()

def primera_mayuscula(valor):
    if not isinstance(valor, str):
        return valor
    return valor.strip().capitalize()

def validar_aproximado(valor, opciones):
    val_norm = normalizar(valor)
    opciones_norm = [normalizar(o) for o in opciones]
    if val_norm in opciones_norm:
        return opciones[opciones_norm.index(val_norm)], True, ""
    match = get_close_matches(val_norm, opciones_norm, n=1, cutoff=0.7)
    if match:
        return opciones[opciones_norm.index(match[0])], True, ""
    return valor, False, "Valor no válido"

def validar_entero(valor):
    try:
        return int(float(valor)), True, ""
    except:
        return valor, False, "Número entero inválido"

def validar_decimal(valor):
    try:
        return round(float(valor), 1), True, ""
    except:
        return valor, False, "Número decimal inválido"

def validar_fecha(valor):
    try:
        fecha = pd.to_datetime(valor, dayfirst=True, errors='raise')
        return fecha.strftime("%d/%m/%Y"), True, ""
    except Exception as e:
        return valor, False, f"Fecha inválida: {e}"

# ---------------- Configuración ----------------
valores_validos = {
    "combustible": ["Nafta", "Diesel", "Gas", "Electrico"],
    "med. uso": ["Kilometros", "Millas", "Horas"],
    "estado": ["Asignado", "Disponible", "En Taller", "Fuera de Servicio"],
    "tipo cobertura": ["Tercero Completo Estandard", "Tercero Completo Premium", "Todo Riesgo"],
    "titularidad": ["Propio", "Alquilado", "Leasing", "Prendario"]
}
columnas_fecha = [
    "vto póliza", "vto cédula", "vto vtv", "vto ruta",
    "vto gnc", "vto cilindro gnc", "vto senasa"
]

# ---------------- App Streamlit ----------------
st.title("Validador de Archivo Excel - Vehículos")

file = st.file_uploader("Subí tu archivo de Excel para validar y corregir los datos automáticamente.", type=[".xlsx"])

if file:
    wb = load_workbook(file)
    ws = wb.active

    encabezado = [str(c.value).strip() if c.value else "" for c in ws[6]]
    encabezado_normalizado = [normalizar_columna(c) for c in encabezado]
    errores = []
    corregidos = []
    cambios_por_columna = {}

    fill_red = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    font_white = Font(color="FFFFFF", bold=True)

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(encabezado)):
        for cell in row:
            col = encabezado_normalizado[cell.column - 1]
            original = cell.value
            nuevo = original
            ok = True
            motivo = ""

            if col == "codigo - interno":
                nuevo = mayusculas(original)
            elif col == "dominio":
                nuevo = limpiar_dominio(original)
            elif col in ["marca", "modelo", "tipo de vehículo", "grupo - base", "cia. seguros", "nombre del titular"]:
                nuevo = titulo_propio(original)
            elif col == "color":
                nuevo = titulo_propio(original)
            elif col in ["nro chasis", "nro motor", "nro póliza"]:
                nuevo = mayusculas(original)
            elif col == "año":
                if isinstance(original, str) and "/" in original:
                    try:
                        nuevo = pd.to_datetime(original, dayfirst=True).year
                    except:
                        pass
                else:
                    nuevo, ok, motivo = validar_entero(original)
            elif col == "cons. promedio":
                nuevo, ok, motivo = validar_decimal(original)
            elif col in valores_validos:
                nuevo, ok, motivo = validar_aproximado(original, valores_validos[col])
            elif col in columnas_fecha:
                nuevo, ok, motivo = validar_fecha(original)
            elif col == "comentarios":
                nuevo = primera_mayuscula(original)

            if not ok:
                cell.fill = fill_red
                cell.font = font_white
                errores.append((cell.row, encabezado[cell.column - 1], original, motivo))
            elif nuevo != original:
                cell.value = nuevo
                corregidos.append((cell.row, encabezado[cell.column - 1], original, nuevo))
                cambios_por_columna[encabezado[cell.column - 1]] = cambios_por_columna.get(encabezado[cell.column - 1], 0) + 1

    st.markdown("### 📊 Resumen general")
    st.info(f"Se detectaron **{len(errores)} errores** y se corrigieron automáticamente **{len(corregidos)} valores**.")

    with st.expander("📋 Ver resumen de cambios detectados"):
        if errores:
            st.markdown("### ⚠️ Cambios no corregidos")
            st.dataframe(pd.DataFrame(errores, columns=["Fila", "Columna", "Valor original", "Observación"]))

        if corregidos:
            st.markdown("### ✅ Cambios realizados")
            st.markdown("📊 Cambios por columna:")
            st.dataframe(pd.DataFrame.from_dict(cambios_por_columna, orient="index", columns=["Cantidad de cambios"]))
            st.markdown("📝 Detalle de cambios:")
            st.dataframe(pd.DataFrame(corregidos, columns=["Fila", "Columna", "Valor original", "Valor corregido"]))

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.download_button(
        label="📥 Descargar archivo validado",
        data=output,
        file_name="vehiculos_validado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
