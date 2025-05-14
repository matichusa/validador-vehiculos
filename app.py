
# Simulación base para validador con control de duplicados

# ... (imports y configuración arriba)

valores_unicos = {
    "dominio": set()
}

# ... (inicio del for row/cell)


# Este fragmento reemplaza la lógica del campo dominio en el validador

if col_name == "dominio" and isinstance(val, str):
    nuevo = limpiar_dominio(val)
    if nuevo in valores_unicos["dominio"]:
        cell.fill = fill_red
        cell.font = font_white
        errores.append((cell.row, encabezados[cell.column - 1], val, "Duplicado"))
    else:
        valores_unicos["dominio"].add(nuevo)
        if nuevo != val:
            corregidos.append((cell.row, encabezados[cell.column - 1], val, nuevo))
            cambios_por_columna[encabezados[cell.column - 1]] = cambios_por_columna.get(encabezados[cell.column - 1], 0) + 1
            cell.value = nuevo


# ... (exportación y visualización)
