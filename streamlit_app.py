import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Limpiar Excel - Ingresos/Egresos")
st.title("📊 Limpiar archivo Excel de Ingresos/Egresos")
st.write("Subí tu archivo original para generar uno limpio, con las columnas necesarias.")

archivo = st.file_uploader("📤 Subí el archivo original Excel", type=[".xlsx"])

columnas_a_conservar = [
    "Guia/PLAN", "Origen", "Destino", "Empresa",
    "Identificador", "Nombre/Descripcion", "Proyecto"
]

if archivo:
    df = pd.read_excel(archivo)

    # Filtramos solo las columnas necesarias
    df_limpio = df[columnas_a_conservar]

    # Crear archivo Excel en memoria
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_limpio.to_excel(writer, index=False)
    output.seek(0)

    # Generar nombre con fecha actual
    fecha_actual = datetime.now().strftime("%d-%m-%Y")
    nombre_archivo = f"INGRESOS-EGRESOS {fecha_actual}.xlsx"

    st.success("✅ Archivo procesado correctamente. Podés descargarlo abajo.")
    st.download_button(
        label="📥 Descargar archivo limpio",
        data=output,
        file_name=nombre_archivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
  
