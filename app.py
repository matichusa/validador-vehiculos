# Validador Excel Vehículos - Restaurado
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from io import BytesIO
from difflib import get_close_matches
import unicodedata

# Normalización
def normalizar_columna(nombre):
    if not isinstance(nombre, str):
        return ""
    nombre = unicodedata.normalize("NFKD", nombre).encode("ascii", "ignore").decode("utf-8")
    return nombre.strip().lower()

def titulo_propio(valor):
    if not isinstance(valor, str):
        return valor
    return " ".join(p.capitalize() for p in valor.strip().split())

def mayusculas(valor):
    if not isinstance(valor, str):
        return valor
    return valor.strip().upper()

def limpiar_dominio(valor):
    if not isinstance(valor, str):
        return valor
    return ''.join(c for c in valor.upper() if c.isalnum())

def validar_aproximado(valor, opciones):
    val_norm = normalizar_columna(valor)
    opciones_norm = [normalizar_columna(o) for o in opciones]
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
        if pd.isna(valor) or str(valor).strip() == "":
            return "", True, ""

        if isinstance(valor, str):
            val = valor.strip().lower().replace(",", "").replace("del ", "").replace(" de ", " ")
            if "00:00:00" in val:
                try:
                    fecha = pd.to_datetime(val.split()[0], format="%Y-%m-%d", errors='raise')
                    return fecha.strftime("%d/%m/%Y"), True, ""
                except:
                    pass
            meses = {
                "enero": "01", "febrero": "02", "marzo": "03", "abril": "04",
                "mayo": "05", "junio": "06", "julio": "07", "agosto": "08",
                "septiembre": "09", "setiembre": "09", "octubre": "10",
                "noviembre": "11", "diciembre": "12"
            }
            for mes, num in meses.items():
                if mes in val:
                    val = val.replace(mes, num)
                    break
            val = val.replace("-", "/").replace(".", "/")
            partes = val.split()
            if len(partes) == 3 and all(any(c.isdigit() for c in p) for p in partes):
                val = "/".join(partes)
            fecha = pd.to_datetime(val, dayfirst=True, errors='raise')
            return fecha.strftime("%d/%m/%Y"), True, ""

        if isinstance(valor, (pd.Timestamp, datetime)):
            return pd.to_datetime(valor).strftime("%d/%m/%Y"), True, ""

        fecha = pd.to_datetime(valor, dayfirst=True, errors='raise')
        return fecha.strftime("%d/%m/%Y"), True, ""

    except Exception as e:
        return valor, False, f"Fecha inválida: {e}"
    try:
        if pd.isna(valor) or str(valor).strip() == "":
            return "", True, ""

        if isinstance(valor, (pd.Timestamp, datetime)):
            return pd.to_datetime(valor).strftime("%d/%m/%Y"), True, ""

        if isinstance(valor, str):
            val = valor.lower().strip()
            val = val.replace(",", "").replace("del ", "").replace(" de ", " ")
            meses = {
                "enero": "01", "febrero": "02", "marzo": "03", "abril": "04",
                "mayo": "05", "junio": "06", "julio": "07", "agosto": "08",
                "septiembre": "09", "setiembre": "09", "octubre": "10",
                "noviembre": "11", "diciembre": "12"
            }
            for mes_texto, mes_num in meses.items():
                if mes_texto in val:
                    val = val.replace(mes_texto, mes_num)
                    break

            val = val.replace("-", "/").replace(".", "/")
            partes = val.split()

            if len(partes) == 3 and all(any(c.isdigit() for c in p) for p in partes):
                val = "/".join(partes)

            fecha = pd.to_datetime(val, dayfirst=True, errors='raise')
            return fecha.strftime("%d/%m/%Y"), True, ""

        fecha = pd.to_datetime(valor, dayfirst=True, errors='raise')
        return fecha.strftime("%d/%m/%Y"), True, ""

    except Exception as e:
        return valor, False, f"Fecha inválida: {e}"
    try:
        if pd.isna(valor) or str(valor).strip() == "":
            return "", True, ""

        if isinstance(valor, (pd.Timestamp, datetime)):
            return pd.to_datetime(valor).strftime("%d/%m/%Y"), True, ""

        if isinstance(valor, str):
            val = valor.lower().strip()
            val = val.replace(",", "").replace("del ", "").replace(" de ", " ")
            meses = {
                "enero": "01", "febrero": "02", "marzo": "03", "abril": "04",
                "mayo": "05", "junio": "06", "julio": "07", "agosto": "08",
                "septiembre": "09", "setiembre": "09", "octubre": "10",
                "noviembre": "11", "diciembre": "12"
            }
            for mes_texto, mes_num in meses.items():
                if mes_texto in val:
                    val = val.replace(mes_texto, mes_num)
                    break

            val = val.replace("-", "/").replace(".", "/")
            partes = val.split()

            if len(partes) == 3 and all(any(c.isdigit() for c in p) for p in partes):
                val = "/".join(partes)

            fecha = pd.to_datetime(val, dayfirst=True, errors='raise')
            return fecha.strftime("%d/%m/%Y"), True, ""

        fecha = pd.to_datetime(valor, dayfirst=True, errors='raise')
        return fecha.strftime("%d/%m/%Y"), True, ""

    except Exception as e:
        return valor, False, f"Fecha inválida: {e}"
    return validar_fecha_robusta(valor)

def validar_fecha_robusta(valor):
    try:
        if pd.isna(valor):
            return valor, True, ""
        
        if isinstance(valor, (pd.Timestamp, datetime)):
            return pd.to_datetime(valor).strftime("%d/%m/%Y"), True, ""

        if isinstance(valor, str):
            val = valor.lower().strip()
            val = val.replace(",", "").replace("del ", "").replace(" de ", " ")
            meses = {
                "enero": "01", "febrero": "02", "marzo": "03", "abril": "04",
                "mayo": "05", "junio": "06", "julio": "07", "agosto": "08",
                "septiembre": "09", "setiembre": "09", "octubre": "10",
                "noviembre": "11", "diciembre": "12"
            }
            for mes_texto, mes_num in meses.items():
                if mes_texto in val:
                    val = val.replace(mes_texto, mes_num)
                    break

            val = val.replace("-", "/").replace(".", "/")
            partes = val.split()

            if len(partes) == 3 and all(any(c.isdigit() for c in p) for p in partes):
                val = "/".join(partes)

            fecha = pd.to_datetime(val, dayfirst=True, errors='raise')
            return fecha.strftime("%d/%m/%Y"), True, ""
        
        fecha = pd.to_datetime(valor, dayfirst=True, errors='raise')
        return fecha.strftime("%d/%m/%Y"), True, ""

    except Exception as e:
        return valor, False, f"Fecha inválida: {e}"
    return validar_fecha_avanzada(valor)

def validar_fecha_avanzada(valor):
    try:
        fecha = pd.to_datetime(valor, dayfirst=True, errors='raise')
        return fecha.strftime("%d/%m/%Y"), True, ""
    except Exception as e:
        return valor, False, f"Fecha inválida: {e}"

# Configuración
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

st.title("Validador Excel Vehículos")

file = st.file_uploader("Subí el archivo Excel", type=["xlsx"])

if file:
    wb = load_workbook(file)
    ws = wb.active

    encabezado = [str(c.value).strip() if c.value else "" for c in ws[6]]
    encabezado_normalizado = [normalizar_columna(c) for c in encabezado]

    errores = []
    corregidos = []
    cambios_por_columna = {}

    rojo = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    blanco = Font(color="FFFFFF", bold=True)

    for row in ws.iter_rows(min_row=7, max_row=ws.max_row, max_col=len(encabezado)):
        for cell in row:
            col_idx = cell.column - 1
            if col_idx >= len(encabezado):
                continue

            col = encabezado_normalizado[col_idx]
            col_original = encabezado[col_idx]
            original = cell.value

            if original is None or str(original).strip() == "":
                continue

            nuevo = original
            ok = True
            motivo = ""

            if col == "codigo - interno":
                nuevo = mayusculas(original)
            elif col == "dominio":
                nuevo = limpiar_dominio(original)
            elif col in ["marca", "modelo", "tipo de vehiculo", "grupo - base", "cia. seguros", "nombre del titular"]:
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
            elif col == "color":
                nuevo = titulo_propio(original)
            elif col == "cons. promedio":
                nuevo, ok, motivo = validar_decimal(original)
            elif col in valores_validos:
                nuevo, ok, motivo = validar_aproximado(original, valores_validos[col])
            elif col in [normalizar_columna(c) for c in columnas_fecha]:
                nuevo, ok, motivo = validar_fecha(original)
            elif col == "comentarios":
                nuevo = titulo_propio(original.lower())

            if not ok:
                cell.fill = rojo
                cell.font = blanco
                errores.append((cell.row, col_original, original, motivo))
            elif str(nuevo) != str(original):
                cell.value = nuevo
                corregidos.append((cell.row, col_original, original, nuevo))
                cambios_por_columna[col_original] = cambios_por_columna.get(col_original, 0) + 1

    st.markdown("### 📊 Resumen general")
    st.info(f"Se detectaron **{len(errores)} errores** y se corrigieron automáticamente **{len(corregidos)} valores**.")

    with st.expander("📋 Ver resumen de cambios detectados"):
        if errores:
            st.markdown("### ⚠️ Cambios no corregidos")
            st.dataframe(pd.DataFrame(errores, columns=["Fila", "Columna", "Valor original", "Observación"]))
        if corregidos:
            st.markdown("### ✅ Cambios realizados")
            st.dataframe(pd.DataFrame(corregidos, columns=["Fila", "Columna", "Valor original", "Valor corregido"]))

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.download_button("📥 Descargar archivo validado", data=output, file_name="vehiculos_validado.xlsx")
