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
    return valor, False, "Valor no válido"

def validar_entero(valor):
    try:
        return int(float(valor)), True, ""
    except:
        return valor, False, "Número entero inválido"

def validar_fecha(valor):
    try:
        sugerencia = None
        if isinstance(valor, str):
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
                    limpio = valor.replace(mes, num).replace(" ", "/").replace("-", "/").replace(".", "/")
                    partes = limpio.split("/")
                    if len(partes) == 2:
                        partes.insert(0, "01")
                        sugerencia = "/".join(partes)
                        valor = sugerencia
                        break
                    elif len(partes) == 3:
                        valor = "/".join(partes)
                        break
            else:
                if "/" in valor or "-" in valor:
                    partes = valor.replace("-", "/").split("/")
                    if len(partes) == 2:
                        partes.insert(0, "01")
                        sugerencia = "/".join(partes)
                        valor = sugerencia
        fecha = pd.to_datetime(valor, dayfirst=True, errors='raise')
        return fecha.strftime("%d/%m/%Y"), True, ""
    except Exception:
        if sugerencia:
            return valor, False, f"Fecha inválida. ¿Quisiste decir: {sugerencia}?"
        return valor, False, "Fecha inválida"
    except Exception as e:
        return valor, False, f"Fecha inválida: {e}"
    except Exception as e:
        return valor, False, f"Fecha inválida: {e}"

# ---------------- Configuración ----------------
valores_validos = {
    "combustible": ["Nafta", "Diésel", "Gas", "Eléctrico"],
    "med. uso": ["Kilometros", "Millas", "Horas"],
    "estado": ["Asignado", "Disponible", "En Taller", "Fuera de Servicio"],
    "tipo de cobertura": ["Terceros Completo Estandard", "Tercero Completo Premium", "Todo Riesgo"],
    "titularidad": ["Propio", "Alquilado", "Leasing", "Prendario"]
}
columnas_fecha = [c.lower() for c in ["Vto Póliza", "Vto Cédula", "Vto VTV", "Vto Ruta", "Vto GNC", "Vto Cilindro GNC", "Vto Senasa"]]

# ---------------- App Streamlit ----------------
st.title("Validador de Archivo Excel - Estética Original")
file = st.file_uploader("Subí el archivo Excel original", type=[".xlsx"])

if file:
    wb = load_workbook(file)
    ws = wb[wb.sheetnames[0]]

    # Buscar encabezado dinámicamente
    encabezado_fila = None
    for fila in ws.iter_rows(min_row=1, max_row=15):
        valores = [str(cell.value).lower().strip() if cell.value else "" for cell in fila]
        if "dominio" in valores:
            encabezado_fila = fila[0].row
            break

    if encabezado_fila is None:
        st.error("No se encontró una fila con encabezados válidos (como 'Dominio').")
        st.stop()

    encabezados = [str(cell.value).strip() if cell.value else "" for cell in ws[encabezado_fila - 1]]
    encabezados_normalizados = [e.lower().strip() for e in encabezados]
    col_map = {col: idx for idx, col in enumerate(encabezados_normalizados)}

    errores = []
    corregidos = []
    cambios_por_columna = {}
    fill_red = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    font_white = Font(color="FFFFFF", bold=True)

    for row in ws.iter_rows(min_row=encabezado_fila + 1, max_row=ws.max_row, max_col=len(encabezados)):
        for cell in row:
            col_idx = cell.column - 1
            if col_idx >= len(encabezados_normalizados):
                continue
            col_name = encabezados_normalizados[col_idx]
            val = cell.value
            original_val = val

            if col_name == "dominio" and isinstance(val, str):
                nuevo = limpiar_dominio(val)
                if nuevo != val:
                    corregidos.append((cell.row, encabezados[cell.column - 1], val, nuevo))
                    cambios_por_columna[encabezados[cell.column - 1]] = cambios_por_columna.get(encabezados[cell.column - 1], 0) + 1
                    cell.value = nuevo

            else:
                match_found = False
                for clave_validada in valores_validos:
                    if normalizar(col_name) == normalizar(clave_validada):
                        nuevo, ok, motivo = validar_aproximado(val, valores_validos[clave_validada])
                        if not ok:
                            cell.fill = fill_red
                            cell.font = font_white
                            errores.append((cell.row, encabezados[cell.column - 1], val, motivo))
                        elif nuevo != val:
                            corregidos.append((cell.row, encabezados[cell.column - 1], val, nuevo))
                            cambios_por_columna[encabezados[cell.column - 1]] = cambios_por_columna.get(encabezados[cell.column - 1], 0) + 1
                            cell.value = nuevo
                        match_found = True
                        break

                if not match_found:
                    if any(normalizar(col_name) == normalizar(f) for f in columnas_fecha):
                        nuevo, ok, motivo = validar_fecha(val)
                        if not ok:
                            cell.fill = fill_red
                            cell.font = font_white
                            errores.append((cell.row, encabezados[cell.column - 1], val, motivo))
                        elif nuevo != val:
                            corregidos.append((cell.row, encabezados[cell.column - 1], val, nuevo))
                            cambios_por_columna[encabezados[cell.column - 1]] = cambios_por_columna.get(encabezados[cell.column - 1], 0) + 1
                            cell.value = nuevo

            elif normalizar(col_name) in ["odometro", "odometros"]:
                nuevo, ok, motivo = validar_entero(val)
                if not ok:
                    cell.fill = fill_red
                    cell.font = font_white
                    errores.append((cell.row, encabezados[cell.column - 1], val, motivo))
                elif nuevo != val:
                    corregidos.append((cell.row, encabezados[cell.column - 1], val, nuevo))
                    cambios_por_columna[encabezados[cell.column - 1]] = cambios_por_columna.get(encabezados[cell.column - 1], 0) + 1
                    cell.value = nuevo

    st.markdown(f"### 📊 Resumen general")
    st.info(f"Se detectaron **{len(errores)} errores** y se corrigieron automáticamente **{len(corregidos)} valores**.")

    with st.expander("📋 Ver resumen de cambios detectados"):
        if errores:
            st.markdown("### ⚠️ Cambios no corregidos")
            st.dataframe(pd.DataFrame(errores, columns=["Fila", "Columna", "Valor original", "Observación"]))
        else:
            st.info("No se detectaron errores no corregibles.")

        st.markdown("### ✅ Cambios realizados")
        if corregidos:
            st.markdown(f"🔢 Total de celdas corregidas: **{len(corregidos)}**")
            st.markdown("📊 Cambios por columna:")
            st.dataframe(pd.DataFrame.from_dict(cambios_por_columna, orient="index", columns=["Cantidad de cambios"]))
            st.markdown("📝 Detalle de cambios:")
            st.dataframe(pd.DataFrame(corregidos, columns=["Fila", "Columna", "Valor original", "Valor corregido"]))
        else:
            st.info("No se realizaron correcciones automáticas.")

    if errores:
        log_ws = wb.create_sheet("Log de Errores")
        log_ws.append(["Fila", "Columna", "Valor original", "Observación"])
        for fila, columna, valor, observacion in errores:
            log_ws.append([fila, columna, valor, observacion])

    if corregidos:
        resumen_ws = wb.create_sheet("Resumen de Cambios")
        resumen_ws.append(["Fila", "Columna", "Valor original", "Valor corregido"])
        for fila, columna, original, nuevo in corregidos:
            resumen_ws.append([fila, columna, original, nuevo])

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.download_button(
        label="📥 Descargar archivo validado",
        data=output,
        file_name="vehiculos_validado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.success("Archivo validado conservando su estética. Se agregaron hojas con log de errores y resumen de cambios si se detectaron.")
