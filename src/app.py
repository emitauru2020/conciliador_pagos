import io
import pandas as pd
from openpyxl.styles import PatternFill
import streamlit as st

# Configuración de la página web
st.set_page_config(page_title="Conciliador de Pagos", page_icon="📊", layout="wide")

st.title("📊 Sistema de Conciliación de Pagos")
st.write("Automatiza el cruce de datos entre tus liquidaciones de Financiera y tu Excel del Local.")
st.write("El sistema mantendrá el formato de tu Local, agrupará las filas verdes arriba y añadirá una pestaña con lo pendiente de la Financiera.")

# Definición de las columnas requeridas (estrictamente en minúsculas para el procesamiento interno)
columnas_local_req = ['financiera', 'aut', 'importe']
columnas_financiera_req = ['producto', 'número de autorización', 'importe de transacción']

# ==========================================
# SECCIÓN DE CARGA DE ARCHIVOS Y CONCILIACIÓN
# ==========================================
col1, col2 = st.columns(2)

with col1:
    excel_pdf_subido = st.file_uploader("Sube el Excel de la Financiera", type=["xlsx"])

with col2:
    excel_casero_subido = st.file_uploader("Sube tu Excel del Local", type=["xlsx"])
    
if excel_pdf_subido is not None and excel_casero_subido is not None:
    st.success("¡Ambos archivos cargados! Listo para iniciar el cruce.")
    
    if st.button("Ejecutar Conciliación 🚀"):
        with st.spinner("Buscando coincidencias exactas y ordenando filas..."):
            try:
                # Leer los archivos cargados
                df_del_pdf = pd.read_excel(excel_pdf_subido)
                df_casero_original = pd.read_excel(excel_casero_subido)
                
                # Crear copias de trabajo para la lógica de limpieza y match
                df_pdf_limpio = df_del_pdf.copy()
                df_casero_limpio = df_casero_original.copy()
                
                # Agregar ID de fila único para rastrear la posición original antes de limpiar headers
                df_casero_limpio['_id_fila'] = range(len(df_casero_limpio))
                
                # INMUNIDAD DE MAYÚSCULAS/MINÚSCULAS EN LOS ENCABEZADOS:
                df_casero_limpio.columns = df_casero_limpio.columns.astype(str).str.strip().str.lower()
                df_pdf_limpio.columns = df_pdf_limpio.columns.astype(str).str.strip().str.lower()
                
                # Verificar que existan las columnas correspondientes en cada archivo antes de procesar
                verificacion_local = all(col in df_casero_limpio.columns for col in columnas_local_req)
                verificacion_financiera = all(col in df_pdf_limpio.columns for col in columnas_financiera_req)
                
                if verificacion_local and verificacion_financiera:
                    
                    # 🛠️ 1. MAPEO Y HOMOLOGACIÓN DE "FINANCIERA" EN EL EXCEL DEL LOCAL
                    df_casero_limpio['financiera'] = df_casero_limpio['financiera'].astype(str).str.strip().str.lower()
                    
                    diccionario_homologacion = {
                        'visa pos': 'VISA',
                        'master pos': 'MASTER',
                        'oca pos': 'OCA'
                    }
                    df_casero_limpio['financiera'] = df_casero_limpio['financiera'].map(diccionario_homologacion).fillna(df_casero_limpio['financiera'].str.upper())
                    
                    # 🛠️ 2. LIMPIEZA GENERAL DE DATOS (Columnas del Local)
                    df_casero_limpio['financiera'] = df_casero_limpio['financiera'].astype(str).str.strip().str.upper()
                    df_casero_limpio['aut'] = df_casero_limpio['aut'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
                    df_casero_limpio['importe'] = pd.to_numeric(df_casero_limpio['importe'], errors='coerce')
                    
                    # 🛠️ 3. LIMPIEZA GENERAL DE DATOS (Columnas de la Financiera)
                    df_pdf_limpio['producto'] = df_pdf_limpio['producto'].astype(str).str.strip().str.upper()
                    df_pdf_limpio['número de autorización'] = df_pdf_limpio['número de autorización'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
                    df_pdf_limpio['importe de transacción'] = pd.to_numeric(df_pdf_limpio['importe de transacción'], errors='coerce')
                    # MARCAR COINCIDENCIAS (Agregamos columnas temporales de control)
                    df_pdf_limpio["_existe_en_pdf"] = True
                    df_casero_limpio["_existe_en_local"] = True

                    # Hacemos un merge completo (outer) para identificar qué sobra en cada lado
                    resultado_cruce = pd.merge(
                        df_casero_limpio[
                            columnas_local_req
                            + ["_id_fila", "_existe_en_local"]
                        ],
                        df_pdf_limpio[
                            columnas_financiera_req + ["_existe_en_pdf"]
                        ],
                        left_on=["financiera", "aut", "importe"],
                        right_on=[
                            "producto",
                            "número de autorización",
                            "importe de transacción",
                        ],
                        how="outer",
                    )

                    # Identificar cuáles IDs del local hicieron match
                    ids_con_match = set(
                        resultado_cruce[
                            (resultado_cruce["_existe_en_local"] == True)
                            & (resultado_cruce["_existe_en_pdf"] == True)
                        ]["_id_fila"]
                    )

                    # 🚀 CORRECCIÓN: Usamos un nombre sin guion bajo al inicio para evitar restricciones de Pandas
                    df_casero_original["coincide_match"] = range(
                        len(df_casero_original)
                    )
                    df_casero_original["coincide_match"] = df_casero_original[
                        "coincide_match"
                    ].isin(ids_con_match)

                    # Ordenar para que True (match) quede arriba y False (no match) quede abajo
                    df_casero_ordenado = df_casero_original.sort_values(
                        by="coincide_match", ascending=False
                    )

                    # Extraer las filas de la FINANCIERA que no tuvieron contraparte en el local
                    df_financiera_no_match_limpio = resultado_cruce[
                        (resultado_cruce["_existe_en_pdf"] == True)
                        & (resultado_cruce["_existe_en_local"].isna())
                    ]

                    # Armamos el DataFrame limpio con los sobrantes de la financiera para exportar
                    df_financiera_no_match = pd.DataFrame()
                    df_financiera_no_match["Producto"] = (
                        df_financiera_no_match_limpio["producto"]
                    )
                    df_financiera_no_match["Número de Autorización"] = (
                        df_financiera_no_match_limpio["número de autorización"]
                    )
                    df_financiera_no_match["Importe de Transacción"] = (
                        df_financiera_no_match_limpio["importe de transacción"]
                    )

                    # Calcular estadísticas para las métricas
                    total_matches = len(ids_con_match)
                    sin_pago_local = len(df_casero_original) - total_matches
                    sobrantes_financiera = len(df_financiera_no_match)

                    # Mostrar métricas en pantalla
                    st.divider()
                    metric_col1, metric_col2, metric_col3 = st.columns(3)
                    with metric_col1:
                        st.metric(
                            "Ventas Conciliadas (Agrupadas Arriba)",
                            total_matches,
                        )
                    with metric_col2:
                        st.metric("Ventas Local sin coincidencia", sin_pago_local)
                    with metric_col3:
                        st.metric(
                            "Líneas Financiera sin match", sobrantes_financiera
                        )
                    st.divider()

                    # 🎨 GENERAR EL EXCEL CON ESTILOS DE COLOR Y DOS PESTAÑAS
                    output_reporte = io.BytesIO()
                    with pd.ExcelWriter(
                        output_reporte, engine="openpyxl"
                    ) as writer:
                        # Removemos la columna técnica 'coincide_match' antes de guardar la Pestaña 1 ordenada
                        df_exportar_local = df_casero_ordenado.drop(
                            columns=["coincide_match"]
                        )
                        df_exportar_local.to_excel(
                            writer, index=False, sheet_name="Conciliacion_Local"
                        )

                        # PESTAÑA 2: Escribimos las transacciones sobrantes de la financiera
                        df_financiera_no_match.to_excel(
                            writer,
                            index=False,
                            sheet_name="No_Conciliados_Financiera",
                        )

                        # Accedemos a la primera hoja para aplicar los estilos de openpyxl
                        workbook = writer.book
                        worksheet = writer.sheets["Conciliacion_Local"]

                        # Definimos el color verde claro (Hex: C6EFCE)
                        relleno_verde = PatternFill(
                            start_color="C6EFCE",
                            end_color="C6EFCE",
                            fill_type="solid",
                        )

                        # 🚀 CORRECCIÓN: Acceso seguro al valor booleano usando getattr() o índices de diccionario
                        for idx, fila_data in enumerate(
                            df_casero_ordenado.itertuples()
                        ):
                            if getattr(fila_data, "coincide_match"):
                                fila_excel = (
                                    idx + 2
                                )  # Fila 1 es el encabezado del Excel
                                for col_idx in range(
                                    1, worksheet.max_column + 1
                                ):
                                    worksheet.cell(
                                        row=fila_excel, column=col_idx
                                    ).fill = relleno_verde

                    datos_reporte = output_reporte.getvalue()

                    st.success(
                        "¡Conciliación finalizada! Archivo generado y ordenado con éxito."
                    )
                    st.download_button(
                        label="📥 Descargar Reporte de Conciliación Completo",
                        data=datos_reporte,
                        file_name="reporte_conciliacion_final.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                else:
                    st.error(
                        f"Error de formato:\n"
                        f"- El Excel del Local debe contener las columnas: financiera, aut, importe.\n"
                        f"- El Excel de la Financiera debe contener las columnas: producto, número de autorización, importe de transacción."
                    )
            except Exception as e:
                st.error(f"Error en la conciliación: {e}")
