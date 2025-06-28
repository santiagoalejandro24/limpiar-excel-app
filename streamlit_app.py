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
    df_limpio = df_limpio.sort_values(by=['Empresa', 'Nombre/Descripcion'], ascending=[True, False])

    # --- ¡NUEVAS LÍNEAS PARA ORDENAR! ---
    # RECORDÁ: Necesitas decidir qué columna(s) usar para ordenar.
    #
    # Ejemplo:
    # Si quieres ordenar por 'Empresa' alfabéticamente (A-Z)
    # y luego por 'Nombre/Descripcion' de mayor a menor (Z-A), usa:
    #
    # df_limpio = df_limpio.sort_values(by=['Empresa', 'Nombre/Descripcion'], ascending=[True, False])
    #
    # Si solo quieres ordenar por una columna, por ejemplo, 'Proyecto' de A-Z:
    # df_limpio = df_limpio.sort_values(by='Proyecto', ascending=True)
    #
    # Si quieres 'Proyecto' de Z-A (mayor a menor para texto):
    # df_limpio = df_limpio.sort_values(by='Proyecto', ascending=False)
    
    # *** ¡Ajusta las columnas y el orden (True para A-Z, False para Z-A) según tus necesidades! ***
    # Aquí te dejo un ejemplo con 'Empresa' (A-Z) y 'Nombre/Descripcion' (Z-A):
    df_limpio = df_limpio.sort_values(by=['Empresa', 'Nombre/Descripcion'], ascending=[True, False])
    st.info("¡Se aplicó un ordenamiento! Asegúrate de que las columnas usadas ('Empresa', 'Nombre/Descripcion' en este ejemplo) sean las que necesitas. ¡Podés cambiarlas en el código!")
    # -----------------------------------

    # Crear archivo Excel en memoria
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_limpio.to_excel(writer, index=False)
    output.seek(0)

    # Generar nombre con fecha actual
    fecha_actual = datetime.now().strftime("%d-%m-%Y")
    nombre_archivo = f"INGRESOS-EGRESOS {fecha_actual}.xlsx"

    st.success("✅ Archivo procesado y ordenado correctamente. ¡Podés descargarlo abajo!")
    st.download_button(
        label="📥 Descargar archivo limpio",
        data=output,
        file_name=nombre_archivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

