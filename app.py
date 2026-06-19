import streamlit as st
import pandas as pd
import io

# Configuración de la página
st.set_page_config(page_title="Super Buscador de Alumnos", page_icon="🎓", layout="wide")

st.title("🎓 Sistema Avanzado de Búsqueda y Filtros")
st.write("Subí el reporte de la web para desglosar la información, aplicar múltiples filtros y exportar.")

# Componente para subir el archivo Excel
archivo_subido = st.file_uploader("Arrastrá o seleccioná el archivo Excel (.xlsx)", type=["xlsx"])

if archivo_subido is not None:
    try:
        # 1. LEER ARCHIVO FILA POR FILA
        df_raw = pd.read_excel(archivo_subido, header=None)
        
        datos_processed = []
        curso_actual = "Sin Curso Detectado"
        columnas_reales = []
        
        for index, fila in df_raw.iterrows():
            valores_fila = fila.dropna().tolist()
            if not valores_fila:
                continue
            
            texto_unido = " ".join([str(x).upper() for x in valores_fila])
            
            # Detectamos la fila de encabezados reales (Alumno, Identificación, etc.)
            if "ALUMNO" in texto_unido or "IDENTIFICACION" in texto_unido or "LEGAJO" in texto_unido:
                if not columnas_reales:
                    columnas_reales = [str(x).strip() for x in fila.tolist() if pd.notna(x)]
                continue
            
            # Si es un filtro del inicio (Año Académico, Ubicación, etc.), lo salteamos
            if "AÑO ACADÉMICO" in texto_unido or "PERÍODO LECTIVO" in texto_unido or "UBICACIÓN" in texto_unido:
                continue
                
            # Si tiene un solo elemento de texto y no es lo anterior, es el nombre de un curso/trayecto
            if len(valores_fila) == 1 and isinstance(valores_fila[0], str) and len(valores_fila[0]) > 5:
                curso_actual = valores_fila[0].strip()
                continue
            
            # Si ya tenemos columnas mapeadas y la fila contiene datos de alumno, los guardamos
            if columnas_reales and len(fila) >= len(columnas_reales):
                datos_alumno = fila.iloc[:len(columnas_reales)].tolist()
                datos_processed.append([curso_actual] + datos_alumno)
        
        if datos_processed and columnas_reales:
            # Creamos la tabla limpia de forma interna
            COLUMNA_PRINCIPAL = "Curso / Trayecto"
            columnas_finales = [COLUMNA_PRINCIPAL] + columnas_reales
            df_limpio = pd.DataFrame(datos_processed, columns=columnas_finales)
            
            # Limpiamos filas extras que repitan palabras claves de encabezado
            df_limpio = df_limpio[~df_limpio.astype(str).apply(lambda x: x.str.contains('Alumno|Identificacion', case=False)).any(axis=1)]
            df_limpio = df_limpio.dropna(subset=[columnas_reales[0]])
            
            # --- PANEL DE CONTROL LATERAL (FILTROS) ---
            st.sidebar.header("🛠️ Panel de Filtros")
            
            # Filtro principal: El Curso
            lista_cursos = sorted(df_limpio[COLUMNA_PRINCIPAL].unique())
            curso_seleccionado = st.sidebar.selectbox("1. Seleccioná el Curso:", lista_cursos)
            
            # Filtramos inicialmente por ese curso
            df_filtrado = df_limpio[df_limpio[COLUMNA_PRINCIPAL] == curso_seleccionado].copy()
            
            # Filtros dinámicos adicionales para el resto de las columnas (Estado, Comisión, etc.)
            st.sidebar.markdown("---")
            st.sidebar.subheader("2. Filtrar por otras columnas:")
            
            for col in columnas_reales:
                valores_unicos = sorted(df_filtrado[col].dropna().astype(str).unique())
                if valores_unicos:
                    # CORREGIDO: Usamos 'options' en inglés para evitar el error previo
                    seleccion = st.sidebar.multiselect(f"Filtrar por {col.title()}:", options=valores_unicos)
                    if seleccion:
                        df_filtrado = df_filtrado[df_filtrado[col].astype(str).isin(seleccion)]
            
            # --- SELECCIÓN DE COLUMNAS VISIBLES ---
            st.sidebar.markdown("---")
            st.sidebar.subheader("3. Columnas Visibles:")
            columnas_visibles = st.sidebar.multiselect(
                "Elegí qué columnas querés ver:",
                options=columnas_reales,
                default=columnas_reales
            )
            
            # --- DESPLIEGUE DE RESULTADOS ---
            st.subheader(f"📊 Resultados para: {curso_seleccionado}")
            
            if not columnas_visibles:
                st.warning("Por favor, seleccioná al menos una columna en el panel izquierdo para visualizar los datos.")
            else:
                # Recortamos la tabla a lo seleccionado por el usuario
                df_final_mostrar = df_filtrado[columnas_visibles]
                
                # Métrica de cantidad
                st.metric(label="Alumnos encontrados", value=len(df_final_mostrar))
                
                # Tabla Interactiva en pantalla
                st.dataframe(df_final_mostrar, use_container_width=True)
                
                # --- SECCIÓN DE DESCARGAS ---
                st.markdown("### 📥 Exportar Resultados")
                col_btn1, col_btn2 = st.columns(2)
                
                # Descarga en CSV
                csv_data = df_final_mostrar.to_csv(index=False).encode('utf-8-sig')
                col_btn1.download_button(
                    label="📄 Descargar en formato CSV",
                    data=csv_data,
                    file_name=f"Filtro_{curso_seleccionado.replace(' ', '_')}.csv",
                    mime="text/csv",
                    use_container_width=True
                )
                
                # Descarga en EXCEL
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_final_mostrar.to_excel(writer, index=False, sheet_name='Datos Filtrados')
                excel_data = buffer.getvalue()
                
                col_btn2.download_button(
                    label="🟢 Descargar en formato EXCEL",
                    data=excel_data,
                    file_name=f"Filtro_{curso_seleccionado.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
        else:
            st.error("No se pudo estructurar el contenido del Excel. Asegurate de que el archivo tenga datos válidos.")
            
    except Exception as e:
        st.error(f"Ocurrió un error inesperado: {e}")