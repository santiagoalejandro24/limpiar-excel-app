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

    # --- ¡AYUDA CRÍTICA PARA DEPURAR! ---
    st.write("---")
    st.write("### 🔍 **¡PASO CLAVE: VERIFICÁ LOS NOMBRES DE TUS COLUMNAS!**")
    st.write("Estas son las columnas **EXACTAS** que tu aplicación detecta en tu **archivo Excel ORIGINAL**:")
    st.code(df.columns.tolist()) # Esto te mostrará la lista de nombres de columnas
    st.write("---")
    # ------------------------------------

    # Filtramos solo las columnas necesarias
    df_limpio = df[columnas_a_conservar]

    # --- VERIFICACIÓN DE COLUMNAS SELECCIONADAS ---
    st.write("### 📋 **COLUMNAS SELECCIONADAS PARA EL PROCESAMIENTO:**")
    st.write("Estas son las columnas que **quedaron después de la limpieza inicial** y están listas para ordenar:")
    st.code(df_limpio.columns.tolist()) # Esto te mostrará las columnas finales
    st.write("---")
    # ---------------------------------------------

    # --- BLOQUE DE ORDENAMIENTO (CON MEJOR CONTROL DE ERRORES) ---
    try:
        # Aquí está la línea que ordena.
        # Es FUNDAMENTAL que 'Nombre/Descripcion' COINCIDA EXACTAMENTE
        # con uno de los nombres que viste en los listados de arriba.
        df_limpio = df_limpio.sort_values(by='Nombre/Descripcion', ascending=False)
        st.info("✅ ¡El archivo ha sido ordenado por 'Nombre/Descripcion' de mayor a menor (Z-A)! 🎉")
    except KeyError as e:
        st.error(f"❌ ¡ERROR FATAL al intentar ordenar! La columna **'Nombre/Descripcion'** NO fue encontrada.")
        st.error(f"Detalle del error: **'{e}'**")
        st.warning("👉 Por favor, revisa los listados de **'Columnas detectadas...'** y **'Columnas seleccionadas...'** arriba.")
        st.warning("Asegúrate de que el nombre de la columna en el código (`'Nombre/Descripcion'`) sea **IDÉNTICO** al que aparece en tu Excel (mayúsculas, minúsculas, espacios, tildes, ¡todo cuenta!).")
        st.warning("¡La descarga del archivo continuará, pero sin el ordenamiento aplicado!")
    # -------------------------------------------------------------

    # Crear archivo Excel en memoria
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_limpio.to_excel(writer, index=False)
    output.seek(0)

    # Generar nombre con fecha actual
    fecha_actual = datetime.now().strftime("%d-%m-%Y")
    nombre_archivo = f"INGRESOS-EGRESOS {fecha_actual}.xlsx"

    st.success("✅ Archivo procesado correctamente. ¡Podés descargarlo abajo!")
    st.download_button(
        label="📥 Descargar archivo limpio",
        data=output,
        file_name=nombre_archivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
