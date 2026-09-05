import streamlit as st
import pandas as pd
import io
from openpyxl.styles import PatternFill

# Configuración de la página web
st.set_page_config(page_title="Conciliador de Pagos", page_icon="📊", layout="wide")

st.title("📊 Sistema de Conciliación de Pagos")
st.write("Automatiza el cruce de datos entre tus liquidaciones de Financiera y tu Excel del Local.")
st.write("El sistema mantendrá el formato de tu Local y pintará de verde las filas conciliadas con éxito.")

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
        with st.spinner("Buscando coincidencias exactas..."):
            try:
                # Leer los archivos cargados (guardamos una copia limpia del local para el archivo final)
                df_del_pdf = pd.read_excel(excel_pdf_subido)
                df_casero_original = pd.read_excel(excel_casero_subido)
                
                # Crear copias de trabajo para la lógica de limpieza y match
                df_pdf_limpio = df_del_pdf.copy()
                df_casero_limpio = df_casero_original.copy()
                
                # INMUNIDAD DE MAYÚSCULAS/MINÚSCULAS EN LOS ENCABEZADOS:
                # Convertimos todos los nombres de las columnas a minúsculas y limpiamos espacios
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
                    
                    # 💥 MARCAR COINCIDENCIAS (Agregamos una columna temporal de control)
                    df_pdf_limpio['_existe_en_pdf'] = True
                    
                    # Hacemos el merge emparejando las columnas cruzadas correspondientes
                    resultado_cruce = pd.merge(
                        df_casero_limpio[columnas_local_req],
                        df_pdf_limpio[columnas_financiera_req + ['_existe_en_pdf']],
                        left_on=['financiera', 'aut', 'importe'],
                        right_on=['producto', 'número de autorización', 'importe de transacción'],
                        how='left'
                    )
                    
                    # Creamos una máscara booleana: True si matcheó con las columnas de la financiera
                    mascara_match = resultado_cruce['_existe_en_pdf'].fillna(False)
                    
                    # Calcular estadísticas para las métricas
                    total_matches = mascara_match.sum()
                    sin_pago = len(mascara_match) - total_matches
                    
                    # Mostrar métricas en pantalla
                    st.divider()
                    metric_col1, metric_col2 = st.columns(2)
                    with metric_col1:
                        st.metric("Ventas Conciliadas (Pintadas de Verde)", total_matches)
                    with metric_col2:
                        st.metric("Ventas sin coincidencia (Sin color)", sin_pago)
                    st.divider()
                    
                    # 🎨 GENERAR EL EXCEL CON ESTILOS DE COLOR
                    output_reporte = io.BytesIO()
                    with pd.ExcelWriter(output_reporte, engine='openpyxl') as writer:
                        # Escribimos el Excel original del local intacto
                        df_casero_original.to_excel(writer, index=False, sheet_name="Conciliacion_Local")
                        
                        # Accedemos a la hoja para aplicar los estilos de openpyxl
                        workbook = writer.book
                        worksheet = writer.sheets["Conciliacion_Local"]
                        
                        # Definimos el color verde claro (Hex: C6EFCE)
                        relleno_verde = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                        
                        # Recorremos la máscara de resultados para pintar las filas correspondientes
                        for idx, hubo_match in enumerate(mascara_match):
                            if hubo_match:
                                fila_excel = idx + 2
                                for col_idx in range(1, worksheet.max_column + 1):
                                    worksheet.cell(row=fila_excel, column=col_idx).fill = relleno_verde
                                    
                    datos_reporte = output_reporte.getvalue()
                    
                    st.success("¡Conciliación finalizada con éxito y filas coloreadas!")
                    st.download_button(
                        label="📥 Descargar Excel del Local Coloreado",
                        data=datos_reporte,
                        file_name="excel_local_conciliado.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error(f"Error de formato:\n"
                             f"- El Excel del Local debe contener las columnas: financiera, aut, importe.\n"
                             f"- El Excel de la Financiera debe contener las columnas: producto, número de autorización, importe de transacción.")
            except Exception as e:
                st.error(f"Error en la conciliación: {e}")
