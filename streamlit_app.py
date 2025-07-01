import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta

st.set_page_config(page_title="Limpiar Excel - Ingresos/Egresos")
st.title("📊 Limpiar archivo Excel de Ingresos/Egresos")
st.write("Subí tu archivo original para generar uno limpio, separado en hojas de Ingresos, Egresos y un Resumen por país.")

archivo = st.file_uploader("📤 Subí el archivo original Excel", type=[".xlsx"])

orden_columnas = [
    "Guia/PLAN", "Origen", "Destino",
    "Nombre/Descripcion", "Identificador",
    "Empresa",  # ← vamos a asegurarnos que sea la de columna Q
    "Proyecto"
]

if archivo:
    df = pd.read_excel(archivo)

    # Si hay dos columnas "Empresa", nos quedamos con la de la columna Q (normalmente renombrada como Empresa.1)
    if "Empresa.1" in df.columns:
        df.drop(columns=["Empresa"], inplace=True)
        df.rename(columns={"Empresa.1": "Empresa"}, inplace=True)

    faltantes = [col for col in orden_columnas if col not in df.columns]
    if faltantes:
        st.error(f"Faltan columnas en el archivo original: {', '.join(faltantes)}")
        st.stop()

    df_limpio = df[orden_columnas].copy()
    df_limpio = df_limpio[~df_limpio["Identificador"].astype(str).str.contains(r"[A-Za-z]", na=False)]

    df_egresos = df_limpio[
        (df_limpio["Origen"].astype(str).str.contains("Batidero", case=False, na=False)) |
        (df_limpio["Destino"].astype(str).str.contains("Guandacol", case=False, na=False))
    ].copy()

    df_ingresos = df_limpio[
        (~df_limpio["Origen"].astype(str).str.contains("Batidero", case=False, na=False)) &
        (df_limpio["Destino"].astype(str).str.strip().str.lower().isin(["batidero", "la brea"]))
    ].copy()

    df_ingresos = df_ingresos.sort_values(by="Origen", ascending=True)
    df_egresos = df_egresos.sort_values(by="Origen", ascending=True)

    ingresos_chile = df_ingresos[df_ingresos["Origen"].astype(str).str.lower().str.contains("chile")]
    ingresos_arg = df_ingresos[~df_ingresos["Origen"].astype(str).str.lower().str.contains("chile")]

    egresos_chile = df_egresos[df_egresos["Destino"].astype(str).str.lower().str.contains("chile")]
    egresos_arg = df_egresos[~df_egresos["Destino"].astype(str).str.lower().str.contains("chile")]

    df_resumen = pd.DataFrame({
        "País": ["Chile", "Argentina"],
        "Ingresos": [len(ingresos_chile), len(ingresos_arg)],
        "Egresos": [len(egresos_chile), len(egresos_arg)]
    })

    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_ingresos.to_excel(writer, index=False, sheet_name='Ingresos')
        df_egresos.to_excel(writer, index=False, sheet_name='Egresos')
        df_resumen.to_excel(writer, index=False, sheet_name='Resumen')

        workbook = writer.book
        border_format = workbook.add_format({'border': 1})
        header_format = workbook.add_format({
            'bold': True,
            'border': 1,
            'bg_color': '#FFFF00'
        })

        col_widths = {
            "Guia/PLAN": 14,
            "Origen": 14,
            "Destino": 12,
            "Empresa": 39,
            "Identificador": 22,
            "Nombre/Descripcion": 35,
            "Proyecto": 13,
        }

        for sheet_name, df_hoja in [("Ingresos", df_ingresos), ("Egresos", df_egresos)]:
            ws = writer.sheets[sheet_name]
            for idx, col in enumerate(df_hoja.columns):
                ws.set_column(idx, idx, col_widths.get(col, 20))
                ws.write(0, idx, col, header_format)
            for i in range(len(df_hoja)):
                for j in range(len(df_hoja.columns)):
                    ws.write(i + 1, j, df_hoja.iat[i, j], border_format)

        ws_resumen = writer.sheets["Resumen"]
        ws_resumen.set_column(0, 0, 14)
        ws_resumen.set_column(1, 2, 15)
        for col_idx, col in enumerate(df_resumen.columns):
            ws_resumen.write(0, col_idx, col, header_format)
        for row_idx in range(len(df_resumen)):
            for col_idx in range(len(df_resumen.columns)):
                ws_resumen.write(row_idx + 1, col_idx, df_resumen.iat[row_idx, col_idx], border_format)

    output.seek(0)

    # Fecha del día siguiente para el nombre del archivo
    fecha_siguiente = (datetime.now() + timedelta(days=1)).strftime("%d-%m-%Y")
    nombre_archivo = f"INGRESOS-EGRESOS {fecha_siguiente}.xlsx"

    st.success("✅ Archivo procesado correctamente. Podés descargarlo abajo.")
    st.download_button(
        label="📥 Descargar archivo con Ingresos, Egresos y Resumen",
        data=output,
        file_name=nombre_archivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
