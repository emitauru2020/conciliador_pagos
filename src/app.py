import io
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import streamlit as st

# Configuración de la página web
st.set_page_config(page_title="Conciliador de Pagos", page_icon="📊", layout="wide")

st.title("📊 Sistema de Conciliación de Pagos")
st.write("Automatiza el cruce de datos entre tus liquidaciones de Financiera y tu Excel del Local.")
st.write("El sistema mantendrá el formato de tu Local, analizando strictly la columna B para la Financiera y la Columna L para matches previos.")

# Definición de las columnas requeridas para la Financiera (en minúsculas)
columnas_financiera_req = ['producto', 'número de autorización', 'importe de transacción']

# Función optimizada para homologar marcas basándose en cómo empiezan las palabras
def homologar_marca_inicial(texto):
    texto_min = str(texto).strip().lower()
    if texto_min.startswith('visa'):
        return 'VISA'
    elif texto_min.startswith('master'):
        return 'MASTER'
    elif texto_min.startswith('oca'):
        return 'OCA'
    return texto_min.upper()

# ==========================================
# SECCIÓN DE CARGA DE ARCHIVOS Y CONCILIACIÓN
# ==========================================
col1, col2 = st.columns(2)

with col1:
    excel_casero_subido = st.file_uploader("Sube tu Excel del Local", type=["xlsx"])

with col2:
    excel_pdf_subido = st.file_uploader("Sube el Excel de la Financiera", type=["xlsx"])
    
if excel_pdf_subido is not None and excel_casero_subido is not None:
    st.success("¡Ambos archivos cargados! Listo para iniciar el cruce.")
    
    if st.button("Ejecutar Conciliación 🚀"):
        with st.spinner("Buscando coincidencias exactas..."):
            try:
                # Leer los archivos cargados
                df_del_pdf = pd.read_excel(excel_pdf_subido)
                df_casero_original = pd.read_excel(excel_casero_subido)
                
                # Crear copias de trabajo para la lógica de limpieza y match
                df_pdf_limpio = df_del_pdf.copy()
                df_casero_limpio = df_casero_original.copy()
                
                # Agregar ID de fila único para rastrear la posición original antes de cualquier limpieza
                df_casero_limpio['_id_fila'] = range(len(df_casero_limpio))
                
                # INMUNIDAD DE MAYÚSCULAS/MINÚSCULAS EN LOS ENCABEZADOS DE LA FINANCIERA:
                df_pdf_limpio.columns = df_pdf_limpio.columns.astype(str).str.strip().str.lower()
                
                # Convertimos temporalmente a minúsculas los nombres de las columnas del local para mapear posiciones
                columnas_local_min = [str(c).strip().lower() for c in df_casero_limpio.columns]
                
                # 🎯 REQUERIMIENTO STRICT: Forzar columna B (Índice 1 en Pandas) como la Financiera del Local
                posicion_columna_b = 1 
                
                # Validar la existencia de las columnas obligatorias 'aut' e 'importe' en el Local
                if 'aut' in columnas_local_min and 'importe' in columnas_local_min:
                    idx_aut = columnas_local_min.index('aut')
                    idx_importe = columnas_local_min.index('importe')
                    
                    # Verificar si existe la Columna L (Índice 11 en Pandas) en el Excel cargado
                    tiene_columna_l = df_casero_limpio.shape[1] > 11

                    # Extraemos los datos del local aislando posiciones
                    df_match_local = pd.DataFrame({
                        'financiera': df_casero_limpio.iloc[:, posicion_columna_b],
                        'aut': df_casero_limpio.iloc[:, idx_aut],
                        'importe': df_casero_limpio.iloc[:, idx_importe],
                        'match_previo': df_casero_limpio.iloc[:, 11].astype(str).str.strip().str.lower() if tiene_columna_l else '',
                        '_id_fila': df_casero_limpio['_id_fila']
                    })
                    verificacion_local = True
                else:
                    verificacion_local = False
                
                verificacion_financiera = all(col in df_pdf_limpio.columns for col in columnas_financiera_req)
                
                if verificacion_local and verificacion_financiera:
                    
                    # 🛠️ 1. MAPEO Y HOMOLOGACIÓN DE "FINANCIERA" BASADO EXCLUSIVAMENTE EN LA COLUMNA B
                    df_match_local['financiera'] = df_match_local['financiera'].apply(homologar_marca_inicial)
                    
                    # 🛠️ 2. LIMPIEZA GENERAL DE DATOS (Columnas del Local)
                    df_match_local['aut'] = df_match_local['aut'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
                    df_match_local['importe'] = pd.to_numeric(df_match_local['importe'], errors='coerce')
                    
                    # 🛠️ 3. LIMPIEZA GENERAL DE DATOS (Columnas de la Financiera)
                    df_pdf_limpio['producto'] = df_pdf_limpio['producto'].apply(homologar_marca_inicial)
                    df_pdf_limpio['número de autorización'] = df_pdf_limpio['número de autorización'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
                    df_pdf_limpio['importe de transacción'] = pd.to_numeric(df_pdf_limpio['importe de transacción'], errors='coerce')
                    
                    # Identificar qué financieras están presentes en este Excel de la tarjeta para evaluar solo esas
                    financieras_en_tarjeta = set(df_pdf_limpio['producto'].dropna().unique())
                    
                    # 🔄 FILTRAR REGISTROS DEL LOCAL QUE YA HICIERON MATCH EN ITERACIONES ANTERIORES (Columna L == 'm')
                    df_match_local_para_cruce = df_match_local[df_match_local['match_previo'] != 'm'].copy()

                    # MARCAR COINCIDENCIAS (Agregamos columnas temporales de control)
                    df_pdf_limpio["_existe_en_pdf"] = True
                    df_match_local_para_cruce["_existe_en_local"] = True

                    # Hacemos un merge completo (outer) omitiendo los registros que ya tenían match previo
                    resultado_cruce = pd.merge(
                        df_match_local_para_cruce[['financiera', 'aut', 'importe', '_id_fila', '_existe_en_local']],
                        df_pdf_limpio[columnas_financiera_req + ["_existe_en_pdf"]],
                        left_on=["financiera", "aut", "importe"],
                        right_on=["producto", "número de autorización", "importe de transacción"],
                        how="outer",
                    )

                    # Identificar cuáles IDs del local hicieron match en ESTA iteración
                    ids_con_match = set(
                        resultado_cruce[
                            (resultado_cruce["_existe_en_local"] == True)
                            & (resultado_cruce["_existe_en_pdf"] == True)
                        ]["_id_fila"]
                    )
                    
                    # Identificar cuáles IDs del local NO tienen coincidencia Y ADEMÁS corresponden a la tarjeta evaluada
                    ids_sin_match_filtrados = set(
                        resultado_cruce[
                            (resultado_cruce["_existe_en_local"] == True) & 
                            (resultado_cruce["_existe_en_pdf"].isna()) &
                            (resultado_cruce["financiera"].isin(financieras_en_tarjeta))
                        ]["_id_fila"]
                    )

                    # Extraer las filas de la FINANCIERA que no tuvieron contraparte en el local
                    df_financiera_no_match_limpio = resultado_cruce[
                        (resultado_cruce["_existe_en_pdf"] == True)
                        & (resultado_cruce["_existe_en_local"].isna())
                    ]

                    # Armamos el DataFrame limpio con los sobrantes de la financiera para exportar
                    df_financiera_no_match = pd.DataFrame()
                    df_financiera_no_match["Producto"] = df_financiera_no_match_limpio["producto"]
                    df_financiera_no_match["Número de Autorización"] = df_financiera_no_match_limpio["número de autorización"]
                    df_financiera_no_match["Importe de Transacción"] = df_financiera_no_match_limpio["importe de transacción"]

                    # Calcular estadísticas reales para las métricas
                    total_matches = len(ids_con_match)
                    sin_pago_local = len(ids_sin_match_filtrados)
                    sobrantes_financiera = len(df_financiera_no_match)

                    # Mostrar métricas en pantalla
                    st.divider()
                    metric_col1, metric_col2, metric_col3 = st.columns(3)
                    with metric_col1:
                        st.metric("Ventas Conciliadas (Verdes + 'm')", total_matches)
                    with metric_col2:
                        st.metric("Ventas Local sin coincidencia (Rojas)", sin_pago_local)
                    with metric_col3:
                        st.metric("Líneas Financiera sin match", sobrantes_financiera)
                    st.divider()

                    # 🎨 PROCESAMIENTO VISUAL EXACTO SOBRE EL EXCEL ORIGINAL DEL LOCAL
                    excel_casero_subido.seek(0)
                    wb_local = load_workbook(excel_casero_subido)
                    ws_local = wb_local.active 

                    # Definimos los colores pastel suaves
                    relleno_rojo = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
                    relleno_match = PatternFill(start_color="C5D9F1", end_color="C5D9F1", fill_type="solid")

                    # 1. Marcar con 'm' en la Columna L (Columna 12 en openpyxl) y pintar de VERDE celdas A:D para los nuevos MATCHES
                    for fila_idx in ids_con_match:
                        fila_excel = fila_idx + 2
                        ws_local.cell(row=fila_excel, column=12, value="m")
                        for col_idx in range(1, 5):  
                            ws_local.cell(row=fila_excel, column=col_idx).fill = relleno_match

                    # 2. Pintar de ROJO SOLO hasta la columna D (A, B, C, D) para las filas pendientes de esta financiera
                    for fila_idx in ids_sin_match_filtrados:
                        fila_excel = fila_idx + 2
                        for col_idx in range(1, 5):  
                            ws_local.cell(row=fila_excel, column=col_idx).fill = relleno_rojo

                    # Guardamos el archivo local modificado en un buffer de memoria dedicado
                    output_local = io.BytesIO()
                    wb_local.save(output_local)
                    datos_local_final = output_local.getvalue()

                    # Guardamos los sobrantes de la financiera en su propio archivo Excel independiente
                    output_financiera = io.BytesIO()
                    with pd.ExcelWriter(output_financiera, engine="openpyxl") as writer_fin:
                        df_financiera_no_match.to_excel(writer_fin, index=False, sheet_name="Pendientes_Financiera")
                    datos_financiera_final = output_financiera.getvalue()

                    st.success("¡Conciliación finalizada con éxito! Archivos listos para descargar.")
                    
                    # Botones de descarga organizados de forma limpia en dos columnas
                    down_col1, down_col2 = st.columns(2)
                    with down_col1:
                        st.download_button(
                            label="📥 Descargar tu Excel del Local (Actualizado)",
                            data=datos_local_final,
                            file_name="local_conciliado_alertas.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                        st.caption("Añade 'm' en la Columna L y pinta de verde (coincidencias) o rojo (pendientes) las columnas A-D.")

                    with down_col2:
                        st.download_button(
                            label="📥 Descargar Sobrantes de la Financiera",
                            data=datos_financiera_final,
                            file_name="financiera_sin_local.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                        st.caption("Registros que vinieron en la liquidación pero que no tenías anotados en el Local.")
                else:
                    st.error(
                        f"Error de formato:\n"
                        f"- El Excel del Local debe contener las columnas obligatorias: 'aut' e 'importe'.\n"
                        f"- El Excel de la Financiera debe contener las columnas: producto, número de autorización, importe de transacción."
                    )
            except Exception as e:
                st.error(f"Error en la conciliación: {e}")