import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

# Configuración de la página
st.set_page_config(page_title="Limpiar Excel - Ingresos/Egresos")

# Título y descripción
st.title("📊 Limpiar archivo Excel de Ingresos/Egresos")
st.write("Subí tu archivo original para generar uno limpio, separado en hojas de Ingresos y Egresos.")

# Subida del archivo
archivo = st.file_uploader("📤 Subí el archivo original Excel", type=[".xlsx"])

# Columnas que se desean conservar
columnas_a_conservar = [
    "Guia/PLAN", "Origen", "Destino", "Empresa", 
    "Identificador", "Nombre/Descripcion", "Proyecto"
]

if archivo:
    # Leemos el archivo
    df = pd.read_excel(archivo)

    # Validar columnas
    faltantes = [col for col in columnas_a_conservar if col not in df.columns]
    if faltantes:
        st.error(f"Faltan columnas en el archivo original: {', '.join(faltantes)}")
        st.stop()

    # Filtramos solo las columnas necesarias
    df_limpio = df[columnas_a_conservar].copy()

    # Eliminamos filas donde Identificador contenga letras
    df_limpio = df_limpio[~df_limpio["Identificador"].astype(str).str.contains(r"[A-Za-z]", na=False)]

    # Egresos: Origen contiene "Batidero" o Destino contiene "Guandacol"
    df_egresos = df_limpio[
        (df_limpio["Origen"].astype(str).str.contains("Batidero", case=False, na=False)) |
        (df_limpio["Destino"].astype(str).str.contains("Guandacol", case=False, na=False))
    ].copy()

    # Ingresos: Origen NO contiene "Batidero" y Destino es "Batidero" o "La Brea"
    df_ingresos = df_limpio[
        (~df_limpio["Origen"].astype(str).str.contains("Batidero", case=False, na=False)) &
        (df_limpio["Destino"].astype(str).str.strip().str.lower().isin(["batidero", "la brea"]))
    ].copy()

    # Ordenar por Origen
    df_ingresos = df_ingresos.sort_values(by="Origen", ascending=True)
    df_egresos = df_egresos.sort_values(by="Origen", ascending=True)

    # Crear archivo Excel con dos hojas
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_ingresos.to_excel(writer, index=False, sheet_name='Ingresos')
        df_egresos.to_excel(writer, index=False, sheet_name='Egresos')

        # Ajustar anchos de columnas en ambas hojas
        workbook = writer.book
        for sheet_name, df_hoja in [("Ingresos", df_ingresos), ("Egresos", df_egresos)]:
            worksheet = writer.sheets[sheet_name]
            col_widths = {
                "Guia/PLAN": 14,
                "Origen": 18,
                "Destino": 18,
                "Empresa": 28,
                "Identificador": 22,
                "Nombre/Descripcion": 35,
                "Proyecto": 22,
            }
            for idx, col in enumerate(df_hoja.columns):
                width = col_widths.get(col, 20)
                worksheet.set_column(idx, idx, width)

    output.seek(0)

    # Nombre de archivo con fecha
    fecha_actual = datetime.now().strftime("%d-%m-%Y")
    nombre_archivo = f"INGRESOS-EGRESOS {fecha_actual}.xlsx"

    # Botón de descarga
    st.success("✅ Archivo procesado correctamente. Podés descargarlo abajo.")
    st.download_button(
        label="📥 Descargar archivo con Ingresos y Egresos",
        data=output,
        file_name=nombre_archivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
