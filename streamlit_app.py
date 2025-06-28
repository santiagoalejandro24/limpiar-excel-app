import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

# Configuración de la página
st.set_page_config(page_title="Limpiar Excel - Ingresos/Egresos")

# Título y descripción
st.title("📊 Limpiar archivo Excel de Ingresos/Egresos")
st.write("Subí tu archivo original para generar uno limpio, con las columnas necesarias.")

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

    # Filtramos solo las columnas necesarias
    df_limpio = df[columnas_a_conservar].copy()

    # Intentamos convertir "Identificador" a número
    df_limpio["Identificador_num"] = pd.to_numeric(df_limpio["Identificador"], errors="coerce")

    # Ordenamos: primero los numéricos de mayor a menor, luego los que contienen texto
    df_limpio = df_limpio.sort_values(by="Identificador_num", ascending=False, na_position='last')

    # Quitamos la columna auxiliar
    df_limpio = df_limpio.drop(columns=["Identificador_num"])

    # Crear archivo Excel en memoria
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_limpio.to_excel(writer, index=False)
    output.seek(0)

    # Generar nombre con fecha actual
    fecha_actual = datetime.now().strftime("%d-%m-%Y")
    nombre_archivo = f"INGRESOS-EGRESOS {fecha_actual}.xlsx"

    # Mensaje de éxito y botón de descarga
    st.success("✅ Archivo procesado correctamente. Podés descargarlo abajo.")
    st.download_button(
        label="📥 Descargar archivo limpio",
        data=output,
        file_name=nombre_archivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
